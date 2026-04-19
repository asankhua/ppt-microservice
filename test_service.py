#!/usr/bin/env python3
"""
Test script for PPT Generation Service
Run: python test_service.py
"""

import requests
import json
import sys

BASE_URL = "http://localhost:8000"


def test_health():
    """Test health endpoint"""
    try:
        response = requests.get(f"{BASE_URL}/health")
        if response.status_code == 200:
            print("✅ Health check passed")
            return True
        else:
            print(f"❌ Health check failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Health check error: {e}")
        return False


def test_templates():
    """Test templates endpoint"""
    try:
        response = requests.get(f"{BASE_URL}/templates")
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Templates endpoint: {len(data['templates'])} templates available")
            for t in data['templates']:
                print(f"   - {t['id']}: {t['name']}")
            return True
        else:
            print(f"❌ Templates endpoint failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Templates error: {e}")
        return False


def test_generate():
    """Test PPT generation"""
    payload = {
        "projectName": "Test Food Delivery App",
        "projectDescription": "On-demand food delivery platform for local restaurants",
        "steps": [
            {
                "stepId": 1,
                "stepName": "Problem Reframe",
                "data": {
                    "problemTitle": "High Delivery Costs",
                    "reframedProblem": "How might we reduce delivery costs for small restaurants while maintaining service quality?",
                    "rootCauses": [
                        "Third-party aggregators charge 20-30% commission",
                        "Restaurants have no direct relationship with customers",
                        "Inefficient logistics and route planning"
                    ],
                    "userImpact": "Restaurants lose 25-30% of revenue per order",
                    "opportunitySize": "$50B food delivery market growing 15% annually"
                }
            },
            {
                "stepId": 2,
                "stepName": "Product Vision",
                "data": {
                    "visionStatement": "Empower local restaurants to own their customer relationships and maximize profits through zero-commission online ordering",
                    "elevatorPitch": "DirectBite is a commission-free online ordering platform that helps local restaurants increase profits by 30% while giving them full control of their customer data",
                    "targetAudience": "Independent restaurants with 1-10 locations, annual revenue $500K-$5M",
                    "valueProposition": "Zero commission, direct customer relationships, built-in marketing tools",
                    "successMetrics": [
                        "1,000 restaurants onboarded in Year 1",
                        "$10M GMV processed monthly by Year 2",
                        "30% average revenue increase for restaurants"
                    ]
                }
            },
            {
                "stepId": 3,
                "stepName": "User Personas",
                "data": {
                    "personas": [
                        {
                            "name": "Maria Rodriguez",
                            "role": "Restaurant Owner",
                            "bio": "Owner of 'Maria's Kitchen', a family-owned Mexican restaurant in Chicago with 2 locations. 15 years in business, 25 employees total.",
                            "painPoints": [
                                "Paying 28% commission to delivery apps",
                                "Not knowing who her customers are",
                                "Can't market directly to repeat customers"
                            ],
                            "goals": [
                                "Increase profit margins by 15%",
                                "Build customer email list",
                                "Offer loyalty rewards"
                            ],
                            "techSavviness": "Medium - uses Instagram but needs help with POS integration"
                        },
                        {
                            "name": "James Chen",
                            "role": "Tech-Savvy Owner",
                            "bio": "Owner of 'Noodle House', an Asian fusion restaurant in Seattle. Early adopter of technology, active on social media.",
                            "painPoints": [
                                "Aggregators own his customer relationships",
                                "No control over brand presentation",
                                "Data siloed across multiple platforms"
                            ],
                            "goals": [
                                "Own customer data and analytics",
                                "Integrate with existing systems",
                                "Scale to 5 locations"
                            ],
                            "techSavviness": "High - uses multiple SaaS tools, comfortable with APIs"
                        }
                    ]
                }
            },
            {
                "stepId": 5,
                "stepName": "Market Analysis",
                "data": {
                    "marketOverview": "The online food delivery market is projected to reach $200B globally by 2027, growing at 12% CAGR. However, restaurant profit margins are being squeezed by high aggregator commissions.",
                    "competitors": [
                        {
                            "name": "DoorDash",
                            "description": "Largest US delivery platform",
                            "strengths": [
                                "Massive delivery fleet",
                                "Strong brand recognition",
                                "Advanced logistics technology"
                            ],
                            "weaknesses": [
                                "High commission rates (20-30%)",
                                "Poor restaurant profit margins",
                                "No customer ownership for restaurants"
                            ]
                        },
                        {
                            "name": "Toast",
                            "description": "Restaurant POS with online ordering",
                            "strengths": [
                                "Integrated POS system",
                                "Lower commission rates",
                                "Restaurant-focused features"
                            ],
                            "weaknesses": [
                                "Limited delivery logistics",
                                "Requires Toast POS adoption",
                                "Less consumer brand recognition"
                            ]
                        }
                    ],
                    "tam": "$200B - Global online food delivery market by 2027",
                    "sam": "$25B - US independent restaurant online ordering",
                    "som": "$2.5B - Commission-free platform opportunity (Year 5)"
                }
            },
            {
                "stepId": 7,
                "stepName": "User Stories",
                "data": {
                    "stories": [
                        {
                            "id": "US-001",
                            "asA": "restaurant owner",
                            "iWant": "to set up my online menu in under 30 minutes",
                            "soThen": "I can start accepting orders quickly",
                            "priority": "High",
                            "storyPoints": 5,
                            "riceScore": 85.5
                        },
                        {
                            "id": "US-002",
                            "asA": "customer",
                            "iWant": "to see real-time order status",
                            "soThen": "I know when my food will arrive",
                            "priority": "High",
                            "storyPoints": 3,
                            "riceScore": 72.0
                        },
                        {
                            "id": "US-003",
                            "asA": "restaurant owner",
                            "iWant": "to export my customer email list",
                            "soThen": "I can send promotional emails",
                            "priority": "Medium",
                            "storyPoints": 2,
                            "riceScore": 45.5
                        },
                        {
                            "id": "US-004",
                            "asA": "customer",
                            "iWant": "to save my favorite orders",
                            "soThen": "I can reorder with one click",
                            "priority": "Low",
                            "storyPoints": 2,
                            "riceScore": 32.0
                        }
                    ]
                }
            },
            {
                "stepId": 9,
                "stepName": "OKRs",
                "data": {
                    "northStarDefinition": "Number of commission-free orders processed monthly",
                    "okr1": {
                        "objective": "Achieve product-market fit with independent restaurants",
                        "keyResults": [
                            "Onboard 100 restaurants with 80% retention",
                            "Process 10,000 orders per month",
                            "Achieve 4.5+ star restaurant satisfaction rating"
                        ]
                    },
                    "okr2": {
                        "objective": "Scale to 10 regional markets",
                        "keyResults": [
                            "Expand to 5 major metro areas",
                            "Achieve $1M monthly GMV",
                            "Build 50-person team"
                        ]
                    },
                    "okr3": {
                        "objective": "Become the leading commission-free platform",
                        "keyResults": [
                            "Reach $10M annual revenue",
                            "Process $100M annual GMV",
                            "Achieve 15% market share in target segment"
                        ]
                    }
                }
            }
        ],
        "template": "professional",
        "includeCharts": True,
        "includeImages": False
    }
    
    try:
        print("\n🔄 Generating presentation...")
        response = requests.post(
            f"{BASE_URL}/generate",
            json=payload,
            timeout=60
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print(f"✅ Generation successful: {result['message']}")
                print(f"   File: {result['filePath']}")
                
                # Download the file
                download_url = f"{BASE_URL}{result['downloadUrl']}"
                print(f"\n🔄 Downloading from: {download_url}")
                
                file_response = requests.get(download_url)
                if file_response.status_code == 200:
                    filename = result['downloadUrl'].split('/')[-1]
                    with open(filename, 'wb') as f:
                        f.write(file_response.content)
                    print(f"✅ Downloaded: {filename}")
                    print(f"   Size: {len(file_response.content)} bytes")
                    return True
                else:
                    print(f"❌ Download failed: {file_response.status_code}")
                    return False
            else:
                print(f"❌ Generation failed: {result.get('message')}")
                return False
        else:
            print(f"❌ Generation request failed: {response.status_code}")
            try:
                error = response.json()
                print(f"   Error: {error}")
            except:
                print(f"   Response: {response.text[:200]}")
            return False
    except Exception as e:
        print(f"❌ Generation error: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests"""
    print("=" * 60)
    print("🚀 PPT Service Test Suite")
    print("=" * 60)
    
    results = []
    
    # Test 1: Health
    print("\n1️⃣ Testing Health Endpoint...")
    results.append(("Health", test_health()))
    
    # Test 2: Templates
    print("\n2️⃣ Testing Templates Endpoint...")
    results.append(("Templates", test_templates()))
    
    # Test 3: Generate PPT
    print("\n3️⃣ Testing PPT Generation...")
    results.append(("Generation", test_generate()))
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 Test Summary")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status}: {name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed!")
        return 0
    else:
        print(f"\n⚠️ {total - passed} test(s) failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
