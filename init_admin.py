#!/usr/bin/env python3
"""
Initialize default admin user in MongoDB Manage collection
Run this script to create a default admin user for testing
"""

from pymongo import MongoClient
import sys

def init_admin_user():
    try:
        # Connect to MongoDB
        client = MongoClient('mongodb://127.0.0.1:27017/')
        db = client['local']
        manage_collection = db['Manager']
        
        # Check if admin user already exists
        existing_admin = manage_collection.find_one({'username': 'admin'})
        if existing_admin:
            print("Admin user already exists!")
            print(f"Username: {existing_admin['username']}")
            return
        
        # Create default admin user
        admin_user = {
            'username': 'admin',
            'password': 'admin123',  # Change this in production!
            'role': 'admin',
            'created_at': '2024-01-01',
            'description': 'Default administrator account'
        }
        
        # Insert admin user
        result = manage_collection.insert_one(admin_user)
        
        if result.inserted_id:
            print("✅ Default admin user created successfully!")
            print("Username: admin")
            print("Password: admin123")
            print("⚠️  Please change the password in production!")
        else:
            print("❌ Failed to create admin user")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
    finally:
        if 'client' in locals():
            client.close()

if __name__ == "__main__":
    print("Initializing default admin user...")
    init_admin_user()
