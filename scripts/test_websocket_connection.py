#!/usr/bin/env python3
"""
test_websocket_connection.py

Simple script to test WebSocket server connection.
"""

import asyncio
import json
import websockets
from datetime import datetime

async def test_connection():
    """Test WebSocket connection and basic messaging"""
    uri = "ws://localhost:8081"

    try:
        print(f"Connecting to {uri}...")
        async with websockets.connect(uri) as websocket:
            print("✅ Connected successfully!")

            # Test message
            test_job_id = "test_job_123"

            # Subscribe to a test job
            subscribe_msg = {
                "type": "subscribe",
                "job_id": test_job_id
            }
            await websocket.send(json.dumps(subscribe_msg))
            print(f"📡 Subscribed to job: {test_job_id}")

            # Listen for messages
            print("👂 Listening for messages (5 seconds)...")

            messages_received = 0
            start_time = datetime.now()

            try:
                while (datetime.now() - start_time).total_seconds() < 5:
                    try:
                        message = await asyncio.wait_for(websocket.recv(), timeout=1.0)
                        data = json.loads(message)
                        messages_received += 1
                        print(f"📨 Received: {data.get('type', 'unknown')}")

                        if data.get('type') == 'connection_established':
                            print(f"   Client ID: {data.get('client_id')}")

                    except asyncio.TimeoutError:
                        continue

            except Exception as e:
                print(f"❌ Error receiving messages: {e}")

            # Test ping
            print("\n🏓 Testing ping...")
            ping_msg = {
                "type": "ping",
                "timestamp": datetime.now().isoformat()
            }
            await websocket.send(json.dumps(ping_msg))

            # Wait for pong
            try:
                response = await asyncio.wait_for(websocket.recv(), timeout=2.0)
                pong_data = json.loads(response)
                if pong_data.get('type') == 'pong':
                    print("✅ Ping-pong successful!")
            except asyncio.TimeoutError:
                print("⚠️  No pong response received")

            print(f"\n📊 Test Results:")
            print(f"   Messages received: {messages_received}")
            print(f"   Connection duration: {(datetime.now() - start_time).total_seconds():.1f}s")

            if messages_received > 0:
                print("✅ WebSocket server is working correctly!")
            else:
                print("⚠️  No messages received - server might be idle")

    except websockets.exceptions.ConnectionRefused:
        print("❌ Connection refused - is the server running?")
        print("   Start the server with: python scripts/start_realtime_servers.py")
    except Exception as e:
        print(f"❌ Error: {e}")

async def main():
    """Main test function"""
    print("🔌 WebSocket Connection Test")
    print("=" * 40)

    await test_connection()

    print("\n" + "=" * 40)
    print("Test completed!")

if __name__ == "__main__":
    asyncio.run(main())