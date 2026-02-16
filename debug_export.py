#!/usr/bin/env python3
"""
Debug export functionality
"""

from app import app
import json

def debug_export():
    with app.test_client() as client:
        print("Debugging export functionality...")
        
        # Test 1: Check database status
        print("\n1. Checking database status...")
        response = client.get('/api/dataset/status')
        if response.status_code == 200:
            data = response.json()
            print(f"Database connected: {data.get('success', False)}")
            print(f"Total documents: {data.get('count', 0)}")
        else:
            print("❌ Database status check failed")
            return
        
        # Test 2: Check filter endpoint
        print("\n2. Testing filter endpoint...")
        response = client.get('/api/dataset/filter?goal=Safe')
        if response.status_code == 200:
            data = response.json()
            print(f"Filter successful: {data.get('success', False)}")
            print(f"Records found: {data.get('count', 0)}")
            if data.get('records'):
                print("Sample record fields:", list(data['records'][0].keys())[:5])
        else:
            print("❌ Filter endpoint failed")
            print(f"Response: {response.data.decode('utf-8')}")
        
        # Test 3: Test export with different goals
        print("\n3. Testing export with different goals...")
        test_goals = ["Safe", "Food", "Health", "Diet"]
        
        for goal in test_goals:
            response = client.get(f'/api/dataset/export-filtered?goal={goal}')
            if response.status_code == 200:
                content_type = response.headers.get('content-type', '')
                if 'spreadsheetml' in content_type:
                    print(f"✅ Export successful for '{goal}' - Excel file generated")
                else:
                    try:
                        data = response.json()
                        print(f"❌ Export failed for '{goal}': {data.get('error', 'Unknown error')}")
                    except:
                        print(f"❌ Export failed for '{goal}': {response.data.decode('utf-8')[:100]}")
            else:
                print(f"❌ Export failed for '{goal}': HTTP {response.status_code}")

if __name__ == "__main__":
    debug_export()
