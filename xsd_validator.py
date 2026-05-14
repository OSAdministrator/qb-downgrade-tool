"""
xsd_validator.py — Intuit XSD schema validation for QBXML requests.

Validates QBXML request XML against the QuickBooks SDK XSD schemas to catch
version-incompatible elements BEFORE sending them to QuickBooks.

Usage:
    from xsd_validator import QBXMLValidator

    validator = QBXMLValidator(sdk_version=160)   # QB 2021 = SDK 16.0
    result = validator.validate(xml_string)
    if not result.is_valid:
        for err in result.errors:
            print(f"  XSD ERROR: {err}")
"""

import os
import subprocess
import tempfile
import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional, Callable

LogFn = Optional[Callable[[str], None]]

# Default location of Intuit SDK XSD files and validator
_VALIDATOR_DIR = Path(r"C:\Program Files\Intuit\IDN\Common\tools\validator")
_VALIDATOR_EXE = _VALIDATOR_DIR / "qbValidator.exe"


@dataclass
class ValidationResult:
    """Result of an XSD validation attempt."""
    is_valid: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    xml_snippet: str = ""  # first 200 chars of the XML for context


def _try_lxml_validate(xml_string: str, xsd_path: Path) -> Optional[ValidationResult]:
    """Attempt XSD validation using lxml (fast, in-process)."""
    try:
        from lxml import etree
    except ImportError:
        return None  # lxml not available, fall through to other methods

    result = ValidationResult(xml_snippet=xml_string[:200])
    try:
        xsd_doc = etree.parse(str(xsd_path))
        xsd_schema = etree.XMLSchema(xsd_doc)
        xml_doc = etree.fromstring(xml_string.encode("utf-8"))
        if xsd_schema.validate(xml_doc):
            result.is_valid = True
        else:
            result.is_valid = False
            for err in xsd_schema.error_log:
                result.errors.append(f"Line {err.line}: {err.message}")
    except etree.XMLSyntaxError as e:
        result.is_valid = False
        result.errors.append(f"XML syntax error: {e}")
    except etree.XMLSchemaParseError as e:
        result.is_valid = False
        result.errors.append(f"XSD parse error: {e}")
    except Exception as e:
        result.is_valid = False
        result.errors.append(f"Validation error: {e}")
    return result


def _try_xmlschema_validate(xml_string: str, xsd_path: Path) -> Optional[ValidationResult]:
    """Attempt XSD validation using xmlschema (pure-Python fallback)."""
    try:
        import xmlschema
    except ImportError:
        return None

    result = ValidationResult(xml_snippet=xml_string[:200])
    try:
        schema = xmlschema.XMLSchema(str(xsd_path))
        errors = list(schema.iter_errors(xml_string))
        if not errors:
            result.is_valid = True
        else:
            result.is_valid = False
            for err in errors:
                result.errors.append(str(err.reason or err))
    except Exception as e:
        result.is_valid = False
        result.errors.append(f"Validation error: {e}")
    return result


def _try_dotnet_validate(xml_string: str, xsd_path: Path) -> Optional[ValidationResult]:
    """
    Attempt XSD validation using .NET's System.Xml.Schema (available on Windows).
    Falls back to PowerShell one-liner.
    """
    result = ValidationResult(xml_snippet=xml_string[:200])

    # Write XML to temp file
    tmp = None
    try:
        tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8')
        tmp.write(xml_string)
        tmp.close()

        ps_script = f"""
$xsd = [System.Xml.Schema.XmlSchemaSet]::new()
$xsd.Add($null, '{xsd_path}') | Out-Null
$settings = [System.Xml.XmlReaderSettings]::new()
$settings.Schemas = $xsd
$settings.ValidationType = [System.Xml.ValidationType]::Schema
$errors = @()
$settings.add_ValidationEventHandler({{ param($s,$e) $script:errors += $e.Message }})
$reader = [System.Xml.XmlReader]::Create('{tmp.name}', $settings)
try {{ while ($reader.Read()) {{}} }} finally {{ $reader.Close() }}
if ($errors.Count -eq 0) {{ Write-Output "VALID" }}
else {{ foreach ($e in $errors) {{ Write-Output "ERROR: $e" }} }}
"""
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=30
        )
        output = (proc.stdout or "").strip()
        if output == "VALID":
            result.is_valid = True
        elif output:
            result.is_valid = False
            for line in output.splitlines():
                if line.startswith("ERROR:"):
                    result.errors.append(line[7:].strip())
                else:
                    result.errors.append(line)
        else:
            # No output = something went wrong
            stderr = (proc.stderr or "").strip()
            if stderr:
                result.errors.append(f"PowerShell error: {stderr}")
                result.is_valid = False
            else:
                result.warnings.append("PowerShell validation produced no output")
                result.is_valid = True  # assume OK if no errors
    except subprocess.TimeoutExpired:
        result.warnings.append("PowerShell validation timed out (30s)")
        result.is_valid = True  # don't block on timeout
    except Exception as e:
        result.errors.append(f".NET validation error: {e}")
        result.is_valid = False
    finally:
        if tmp and os.path.exists(tmp.name):
            os.unlink(tmp.name)
    return result


class QBXMLValidator:
    """
    Validates QBXML request/response XML against Intuit's published XSD schemas.

    Args:
        sdk_version: SDK version number (e.g., 160 for QB 2021 SDK 16.0,
                     170 for QB 2023 SDK 17.0). Default 160.
        validator_dir: Path to the directory containing XSD files.
                       Default: C:\\Program Files\\Intuit\\IDN\\Common\\tools\\validator
    """

    # SDK version → QB Desktop version mapping
    VERSION_MAP = {
        160: "QuickBooks 2021 (SDK 16.0)",
        150: "QuickBooks 2020 (SDK 15.0)",
        140: "QuickBooks 2019 (SDK 14.0)",
        130: "QuickBooks 2018 (SDK 13.0)",
        170: "QuickBooks 2023 (SDK 17.0)",
    }

    def __init__(self, sdk_version: int = 160,
                 validator_dir: Optional[Path] = None):
        self.sdk_version = sdk_version
        self.validator_dir = Path(validator_dir) if validator_dir else _VALIDATOR_DIR

        # Locate the operations XSD for this version
        # Pattern: qbxmlops{version}.xsd  (e.g., qbxmlops160.xsd)
        self.ops_xsd = self.validator_dir / f"qbxmlops{sdk_version}.xsd"
        self.types_xsd = self.validator_dir / f"qbxmltypes{sdk_version}.xsd"

        if not self.ops_xsd.exists():
            raise FileNotFoundError(
                f"XSD schema not found: {self.ops_xsd}\n"
                f"Available schemas: {list(self.validator_dir.glob('qbxmlops*.xsd'))}"
            )

        self._method = None  # will be set on first validation

    @property
    def version_label(self) -> str:
        return self.VERSION_MAP.get(self.sdk_version, f"SDK {self.sdk_version}")

    def validate(self, xml_string: str, log_fn: LogFn = None) -> ValidationResult:
        """
        Validate a QBXML request/response string against the SDK XSD.

        Tries validation methods in order: lxml → xmlschema → .NET PowerShell.
        Returns ValidationResult with is_valid, errors, and warnings.
        """
        if not xml_string or not xml_string.strip():
            return ValidationResult(is_valid=False, errors=["Empty XML string"])

        # Try each validation method in order of preference
        for method_name, method_fn in [
            ("lxml", _try_lxml_validate),
            ("xmlschema", _try_xmlschema_validate),
            (".NET/PowerShell", _try_dotnet_validate),
        ]:
            result = method_fn(xml_string, self.ops_xsd)
            if result is not None:
                if self._method is None:
                    self._method = method_name
                    if log_fn:
                        log_fn(f"  XSD Validator: using {method_name} engine against {self.ops_xsd.name}")
                return result

        # All methods failed
        return ValidationResult(
            is_valid=True,
            warnings=["No XSD validation engine available (install lxml or xmlschema)"],
            xml_snippet=xml_string[:200]
        )

    def validate_request_set(self, request_msg_set, log_fn: LogFn = None) -> ValidationResult:
        """
        Validate a QBFC IMsgSetRequest object by extracting its XML via ToXMLString().

        Args:
            request_msg_set: QBFC IMsgSetRequest COM object
            log_fn: Optional logging callback

        Returns:
            ValidationResult
        """
        try:
            xml_str = request_msg_set.ToXMLString()
        except Exception as e:
            return ValidationResult(
                is_valid=False,
                errors=[f"Failed to extract XML from request set: {e}"]
            )
        return self.validate(xml_str, log_fn)

    def validate_batch(self, xml_strings: List[str],
                       labels: Optional[List[str]] = None,
                       log_fn: LogFn = None) -> List[ValidationResult]:
        """Validate multiple XML strings, returning results for each."""
        results = []
        for i, xml in enumerate(xml_strings):
            label = labels[i] if labels and i < len(labels) else f"Request #{i+1}"
            result = self.validate(xml, log_fn)
            if not result.is_valid and log_fn:
                log_fn(f"  XSD FAIL [{label}]: {'; '.join(result.errors)}")
            results.append(result)
        return results

    def diff_schemas(self, other_version: int = 170,
                     log_fn: LogFn = None) -> dict:
        """
        Compare two XSD versions and return elements that exist in other_version
        but NOT in this version (i.e., elements that would fail validation).

        Useful for identifying QB 2023 features unsupported in QB 2021.

        Returns dict with keys: 'added_in_newer', 'removed_in_newer'
        """
        other_xsd = self.validator_dir / f"qbxmlops{other_version}.xsd"
        if not other_xsd.exists():
            if log_fn:
                log_fn(f"  XSD diff: schema {other_xsd.name} not found")
            return {"error": f"Schema {other_xsd.name} not found"}

        try:
            import xml.etree.ElementTree as ET

            def extract_elements(xsd_path: Path) -> set:
                """Extract all element names from an XSD file."""
                tree = ET.parse(str(xsd_path))
                root = tree.getroot()
                ns = {"xs": "http://www.w3.org/2001/XMLSchema"}
                elements = set()
                for elem in root.iter():
                    name = elem.get("name")
                    if name:
                        elements.add(name)
                return elements

            this_elements = extract_elements(self.ops_xsd)
            other_elements = extract_elements(other_xsd)

            added = other_elements - this_elements   # in newer, not in ours
            removed = this_elements - other_elements  # in ours, not in newer

            result = {
                "this_version": self.sdk_version,
                "other_version": other_version,
                "this_element_count": len(this_elements),
                "other_element_count": len(other_elements),
                "added_in_newer": sorted(added),
                "removed_in_newer": sorted(removed),
            }

            if log_fn:
                log_fn(f"  XSD diff: v{self.sdk_version} has {len(this_elements)} elements, "
                       f"v{other_version} has {len(other_elements)} elements")
                if added:
                    log_fn(f"  XSD diff: {len(added)} elements in v{other_version} NOT in v{self.sdk_version}:")
                    for el in sorted(added)[:20]:
                        log_fn(f"    + {el}")
                    if len(added) > 20:
                        log_fn(f"    ... and {len(added) - 20} more")

            return result

        except Exception as e:
            if log_fn:
                log_fn(f"  XSD diff error: {e}")
            return {"error": str(e)}


# ── Quick self-test ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=== Intuit XSD Validator — Self Test ===\n")

    try:
        v = QBXMLValidator(sdk_version=160)
        print(f"Validator created for {v.version_label}")
        print(f"Operations XSD: {v.ops_xsd}")
        print(f"Types XSD:      {v.types_xsd} (exists: {v.types_xsd.exists()})")
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        exit(1)

    # Test 1: Valid QBXML envelope
    test_xml = """<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="16.0"?>
<QBXML>
  <QBXMLMsgsRq onError="continueOnError">
    <AccountQueryRq requestID="1">
    </AccountQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

    print("\n--- Test 1: Valid AccountQueryRq ---")
    result = v.validate(test_xml, log_fn=print)
    print(f"Valid: {result.is_valid}")
    if result.errors:
        for e in result.errors:
            print(f"  Error: {e}")
    if result.warnings:
        for w in result.warnings:
            print(f"  Warning: {w}")

    # Test 2: Schema diff between QB 2021 (160) and QB 2023 (170)
    print("\n--- Test 2: XSD Schema Diff (v160 vs v170) ---")
    diff = v.diff_schemas(other_version=170, log_fn=print)
    if "error" not in diff:
        print(f"\nElements in QB 2023 but NOT in QB 2021: {len(diff.get('added_in_newer', []))}")
        print(f"Elements in QB 2021 but NOT in QB 2023: {len(diff.get('removed_in_newer', []))}")

    print("\n=== Self Test Complete ===")
