#!/usr/bin/env python3
"""
PR Triage Script: Checkout PR, run tests, smoke, style audit, check flakiness.
Outputs Markdown report with pass/fail, top errors, flakiness flags.
"""

import argparse
import subprocess
import os
import sys
import re
from pathlib import Path
from datetime import datetime
import statistics
import shutil

PROJECT_ROOT = Path(__file__).parent.parent

def run_command(cmd, description, cwd=PROJECT_ROOT):
    """Run shell command and return result."""
    print(f"Running: {description}")
    print(f"Command: {' '.join(cmd) if isinstance(cmd, list) else cmd}")
    result = subprocess.run(cmd, shell=isinstance(cmd, str), capture_output=True, text=True, cwd=cwd)
    if result.returncode != 0:
        print(f"Error: {result.stderr}")
    return result

def parse_pytest_output(stdout, stderr):
    """Parse pytest output for passed, failed, and top errors."""
    output = stdout + stderr
    lines = output.splitlines()
    passed = failed = 0
    errors = []
    for line in lines:
        if line.strip().startswith('PASSED'):
            passed += 1
        elif line.strip().startswith('FAILED'):
            failed += 1
            # Extract test name and reason
            match = re.match(r'FAILED (tests/[^:]+):', line)
            if match:
                test_name = match.group(1)
                errors.append(test_name)
    return passed, failed, errors[:3]

def check_flakiness(pytest_cmd, num_runs=3):
    """Run pytest multiple times and check pass rate variance."""
    pass_rates = []
    for i in range(num_runs):
        result = run_command(pytest_cmd, f"Pytest run {i+1}")
        passed, failed, _ = parse_pytest_output(result.stdout, result.stderr)
        total = passed + failed
        rate = passed / total if total > 0 else 0.0
        pass_rates.append(rate)
        print(f"Run {i+1} pass rate: {rate:.2%}")
    if len(pass_rates) < 2:
        return False, pass_rates
    variance = max(pass_rates) - min(pass_rates)
    return variance > 0.1, pass_rates

def run_smoke_test():
    """Run smoke test for DOCX translation."""
    fixture = PROJECT_ROOT / "tests" / "simple_japanese.docx"
    if not fixture.exists():
        print("Warning: Fixture not found, skipping smoke")
        return False, None
    
    output_dir = PROJECT_ROOT / "tmp" / "triage_smoke"
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_docx = output_dir / f"smoke_{timestamp}.docx"
    bilingual_csv = output_dir / f"smoke_{timestamp}_bilingual.csv"
    
    # Run smoke_translate_docx.py with bilingual
    cmd = [
        "python", "scripts/smoke_translate_docx.py",
        "--input", str(fixture),
        "--output", str(output_docx),
        "--csv-report"  # This generates the csv? Wait, from code, --bilingual-csv is needed in translate
    ]
    # Actually, from smoke code, it adds --bilingual-csv to translate cmd if csv_report or audit
    # But --csv-report copies to artifacts, but generates bilingual if needed.
    # To get csv, use --csv-report, it will run with --bilingual-csv
    cmd.append("--csv-report")
    result = run_command(cmd, "DOCX Smoke Test")
    
    smoke_pass = result.returncode == 0 and output_docx.exists()
    csv_path = bilingual_csv if bilingual_csv.exists() else None  # But from smoke, it's in output.parent
    # Adjust: from smoke code, bilingual_path = output_path.parent / f"{output_path.stem}_bilingual.csv"
    actual_csv = output_docx.parent / f"{output_docx.stem}_bilingual.csv"
    if actual_csv.exists():
        csv_path = actual_csv
    else:
        csv_path = None
    
    return smoke_pass, csv_path

def run_style_audit(csv_path):
    """Run style audit on bilingual CSV."""
    if not csv_path or not csv_path.exists():
        return False, []
    
    cmd = ["python", "scripts/audit_style.py", str(csv_path), "--report", str(csv_path.parent / "triage_audit.csv")]
    result = run_command(cmd, "Style Audit")
    audit_pass = result.returncode == 0
    top_issues = []
    if audit_pass:
        # Parse report for top issues? For now, just status
        pass
    else:
        # Could parse stderr for issues
        lines = result.stderr.splitlines()
        top_issues = [line for line in lines[-3:] if "issue" in line.lower()]  # Rough
    
    return audit_pass, top_issues

def generate_md_report(pr_number, pytest_pass, passed, failed, top_pytest_errors, smoke_pass, audit_pass, top_audit_issues, flaky, pass_rates):
    """Generate Markdown report."""
    overall_pass = pytest_pass and smoke_pass and audit_pass and not flaky
    
    md = f"""# PR Triage Report for PR #{pr_number}

Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Pytest Results
- **Status**: {'PASS' if pytest_pass else 'FAIL'}
- **Passed**: {passed}
- **Failed**: {failed}
"""
    if top_pytest_errors:
        md += "## Top 3 Pytest Errors\n"
        for err in top_pytest_errors:
            md += f"- {err}\n"
    else:
        md += "- No errors\n"

    md += f"""
## Smoke Tests (DOCX Adapted)
- **Status**: {'PASS' if smoke_pass else 'FAIL'}

## Style Audit
- **Status**: {'PASS' if audit_pass else 'FAIL'}
"""
    if top_audit_issues:
        md += "## Top 3 Audit Issues\n"
        for issue in top_audit_issues:
            md += f"- {issue}\n"
    else:
        md += "- No issues\n"

    md += f"""
## Flakiness Check
- **Flaky**: {'YES (>10% variance)' if flaky else 'NO'}
- **Pass Rates**: {[f'{r:.2%}' for r in pass_rates]}

## Overall Status
- **{'PASS' if overall_pass else 'FAIL'}**
"""

    report_path = PROJECT_ROOT / f"triage_pr_{pr_number or 'current'}.md"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(md)
    print(f"Report saved to {report_path}")
    return md

def main():
    parser = argparse.ArgumentParser(description="Triage PR with tests, smoke, and audits.")
    parser.add_argument("--pr-number", type=int, help="PR number to checkout (optional for simulation)")
    parser.add_argument("--post-to-gh", action="store_true", help="Post the triage report as a comment on the PR using gh CLI")
    args = parser.parse_args()
    
    pr_number = args.pr_number
    if pr_number:
        checkout_cmd = f"gh pr checkout {pr_number}"
        result = run_command(checkout_cmd, "Checkout PR branch")
        if result.returncode != 0:
            print("Failed to checkout PR, exiting.")
            sys.exit(1)
        print(f"Checked out PR #{pr_number}")
    else:
        print("Simulating on current branch")
    
    # Run pytest
    pytest_cmd = "pytest tests/ -v"
    result_pytest = run_command(pytest_cmd.split(), "Pytest")  # list for split
    passed, failed, top_pytest_errors = parse_pytest_output(result_pytest.stdout, result_pytest.stderr)
    pytest_pass = failed == 0
    
    # Run smoke
    smoke_pass, csv_path = run_smoke_test()
    
    # Run audit
    audit_pass, top_audit_issues = run_style_audit(csv_path)
    
    # Check flakiness
    flaky, pass_rates = check_flakiness(pytest_cmd)
    
    overall_pass = pytest_pass and smoke_pass and audit_pass and not flaky
    
    # Generate report
    md = generate_md_report(pr_number, pytest_pass, passed, failed, top_pytest_errors, smoke_pass, audit_pass, top_audit_issues, flaky, pass_rates)
    print(md)
    
    if args.post_to_gh:
        if not pr_number:
            print("Error: --post-to-gh requires --pr-number")
            sys.exit(1)
        
        report_path = PROJECT_ROOT / f"triage_pr_{pr_number}.md"
        
        summary = f"""## Quick Triage Summary

Overall Status: {'PASS' if overall_pass else 'FAIL'}
"""
        if top_pytest_errors:
            summary += "Top Pytest Errors:\n" + "\n".join(f"- {e}" for e in top_pytest_errors) + "\n\n"
        if top_audit_issues:
            summary += "Top Audit Issues:\n" + "\n".join(f"- {i}" for i in top_audit_issues) + "\n\n"
        
        summary += "@codex-reviewer Please review the triage report below.\n\n---\n"
        
        with open(report_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        full_body = summary + content
        
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(full_body)
        
        if shutil.which("gh") is None:
            print("gh CLI not found, simulating post to PR...")
            print(f"Would run: gh pr comment {pr_number} --body-file {report_path}")
            preview = summary[:200] + "..." if len(summary) > 200 else summary
            print("Summary preview:", preview)
        else:
            cmd = ["gh", "pr", "comment", str(pr_number), "--body-file", str(report_path)]
            result = run_command(cmd, "Posting triage report to PR")
            if result.returncode == 0:
                print(f"Triage report posted successfully to PR #{pr_number}")
            else:
                print(f"Failed to post triage report to PR #{pr_number}")
                if result.stderr:
                    print(result.stderr)
    
    # Exit with overall status
    sys.exit(0 if overall_pass else 1)

if __name__ == "__main__":
    main()
