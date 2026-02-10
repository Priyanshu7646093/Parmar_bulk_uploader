#!/usr/bin/env python
"""Debug script to test question extraction"""

import re
from pathlib import Path

# Copy the OPTION_LABEL_RE from Five_option.py
OPTION_LABEL_RE = re.compile(r"^\s*[\(\[]?(\d{1,2}|[A-Za-z]|[ivxlcdmIVXLCDM]{1,5})[\)\.\]]\s*")

# Test data - you can modify this with your actual question text
test_question = """Q1. What is the capital of France?
A) London
B) Paris
C) Berlin
D) Madrid
E) Rome
Correct Answer: B
Solution: The capital of France is Paris.
"""

print("=" * 60)
print("TESTING QUESTION EXTRACTION")
print("=" * 60)

print("\nTest Question:")
print(test_question)
print("\n" + "=" * 60)

# Test 1: Question number extraction
pattern = r"Q(\d{1,9})\."
match = re.match(pattern, test_question.strip())
if match:
    print(f"✓ Question Number Found: Q{match.group(1)}")
else:
    print("✗ Question Number NOT found")

# Test 2: Option detection
lines = test_question.strip().split('\n')
print(f"\n✓ Total lines: {len(lines)}")

option_count = 0
for i, line in enumerate(lines, 1):
    print(f"  Line {i}: {line}")
    if OPTION_LABEL_RE.match(line):
        option_count += 1
        print(f"    → OPTION DETECTED (Total: {option_count})")

print(f"\n✓ Total options found: {option_count}")

# Test 3: Correct answer extraction
for line in lines:
    if line.lower().startswith("correct answer"):
        print(f"\n✓ Answer line found: {line}")
        answer_text = line.split(":", 1)[-1].strip()
        ans_match = re.search(r"\b([A-Ea-e1-5])\b", answer_text)
        if ans_match:
            print(f"  → Answer extracted: {ans_match.group(1)}")
        break

print("\n" + "=" * 60)
print("If all items show ✓, the extraction should work fine.")
print("If any show ✗, check your question format.")
print("=" * 60)
