#!/usr/bin/env python3
"""Test PPT service locally"""

import requests
import json
import sys

BASE_URL = "http://localhost:8000"

def test_health():
    try:
        r = requests.get(f"{BASE_URL}/health", timeout=5)
        if r.status_code == 200:
            print("✅ Health check passed")
            return True
        print(f"❌ Health failed: {r.status_code}")
        return False
    except Exception as e:
        print(f"❌ Health error: {e}")
        return False

def test_generate():
    payload = {
        "projectName": "Test Product",
        "projectDescription": "A test product",
        "steps": [
            {
                "stepId": 1,
                "stepName": "Problem Reframe",
                "data": {
                    "problemTitle": "Expensive Delivery",
                    "reframedProblem": "How might we reduce costs?",
                    "rootCauses": ["High fees", "No direct connection"]
                }
            },
            {
                "stepId": 2,
                "stepName": "Product Vision",
                "data": {
                    "visionStatement": "Empower restaurants with direct orders",
                    "elevatorPitch": "Zero-commission platform for restaurants"
                }
            },
            {
                "stepId": 3,
                "stepName": "User Personas",
                "data": {
                    "personas": [
                        {
                            "name": "Maria",
                            "role": "Restaurant Owner",
                            "bio": "Owner of family restaurant in Chicago"
                        }
                    ]
                }
            }
        ],
        "template": "professional"
    }
    
    try:
        r = requests.post(f"{BASE_URL}/generate", json=payload, timeout=30)
        if r.status_code == 200:
            result = r.json()
            if result.get('success'):
                print(f"✅ Generate passed: {result.get('message')}")
                print(f"   Download URL: {result.get('downloadUrl')}")
                return True
        print(f"❌ Generate failed: {r.status_code} - {r.text}")
        return False
    except Exception as e:
        print(f"❌ Generate error: {e}")
        return False

def test_download():
    # Test download after generate
    payload = {
        "projectName": "DL Test",
        "steps": [{"stepId": 1, "stepName": "Test", "data": {}}],
        "template": "minimal"
    }
    
    try:
        r = requests.post(f"{BASE_URL}/generate", json=payload, timeout=30)
        if r.status_code == 200:
            result = r.json()
            download_url = result.get('downloadUrl')
            if download_url:
                dl = requests.get(f"{BASE_URL}{download_url}", timeout=10)
                if dl.status_code == 200 and len(dl.content) > 1000:
                    print(f"✅ Download passed: {len(dl.content)} bytes")
                    return True
        print("❌ Download failed")
        return False
    except Exception as e:
        print(f"❌ Download error: {e}")
        return False

if __name__ == "__main__":
    print("Testing PPT Service...")
    print(f"Base URL: {BASE_URL}\n")
    
    results = []
    results.append(test_health())
    results.append(test_generate())
    results.append(test_download())
    
    print(f"\n{'='*40}")
    print(f"Results: {sum(results)}/{len(results)} passed")
    
    if all(results):
        print("✅ All tests passed!")
        sys.exit(0)
    else:
        print("❌ Some tests failed")
        sys.exit(1)
