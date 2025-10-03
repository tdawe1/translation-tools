#!/usr/bin/env python3
"""
demo_realtime_translation.py

Demo script showing how to use real-time progress tracking.
Simulates a translation job with progress updates.
"""

import asyncio
import json
import random
import time
from typing import Dict, Any
from progress_tracker import get_tracker

async def simulate_translation():
    """Simulate a translation job with real-time updates"""

    # Job details
    job_id = "demo_job_001"
    file_name = "sample_presentation.pptx"
    file_size = 5_242_880  # 5MB
    estimated_tokens = 15_000
    estimated_cost = estimated_tokens * 0.00002  # $0.30

    print(f"Starting demo translation job: {job_id}")
    print(f"File: {file_name} ({file_size:,} bytes)")
    print(f"Estimated: {estimated_tokens:,} tokens, ${estimated_cost:.4f}")

    # Start tracking
    tracker = get_tracker()
    await tracker.connect()

    try:
        # Initialize job
        await tracker.start_job(job_id, file_name, file_size, estimated_tokens, estimated_cost)

        # Simulate extraction phase
        print("\n1. Extracting text...")
        await tracker.update_stage("extracting")
        for i in range(10):
            await tracker.update_progress(progress=i*2)
            await asyncio.sleep(0.2)

        # Simulate translation phases
        print("\n2. Translating content...")
        await tracker.update_stage("translating")

        # Simulate batches
        total_batches = 8
        tokens_per_batch = estimated_tokens // total_batches

        for batch in range(1, total_batches + 1):
            print(f"   Processing batch {batch}/{total_batches}...")

            # Update batch progress
            await tracker.update_batch_progress(batch, total_batches)

            # Simulate translation with varying progress
            batch_tokens = tokens_per_batch + random.randint(-500, 500)
            tokens_so_far = batch * tokens_per_batch

            # Incremental progress within batch
            for i in range(10):
                progress = 20 + (batch * 8) + (i * 0.8)
                await tracker.update_progress(
                    progress=progress,
                    tokens_processed=tokens_so_far + (batch_tokens * i // 10),
                    current_cost=(tokens_so_far + (batch_tokens * i // 10)) * 0.00002
                )
                await asyncio.sleep(0.3)

        # Simulate applying translations
        print("\n3. Applying translations...")
        await tracker.update_stage("applying")
        for i in range(10):
            progress = 90 + i
            await tracker.update_progress(progress=progress)
            await asyncio.sleep(0.2)

        # Simulate finalization
        print("\n4. Finalizing document...")
        await tracker.update_stage("finalizing")
        await asyncio.sleep(1)

        # Calculate quality score (random between 0.85 and 0.98)
        quality_score = 0.85 + random.random() * 0.13
        await tracker.set_quality_score(quality_score)

        # Complete job
        final_cost = estimated_tokens * 0.00002
        await tracker.complete_job(success=True)

        print(f"\n✅ Translation completed!")
        print(f"   Final cost: ${final_cost:.4f}")
        print(f"   Quality score: {quality_score*100:.1f}%")

    except Exception as e:
        print(f"\n❌ Translation failed: {e}")
        await tracker.complete_job(success=False, error_message=str(e))

    finally:
        await tracker.disconnect()

async def main():
    """Main demo function"""
    print("🚀 Real-time Translation Progress Demo")
    print("=" * 50)
    print("\nMake sure the WebSocket server is running:")
    print("   python scripts/start_realtime_servers.py")
    print("\nAnd the frontend is running:")
    print("   cd frontend && npm run dev")
    print("\nStarting demo in 3 seconds...")

    await asyncio.sleep(3)

    # Run demo
    await simulate_translation()

    print("\n" + "=" * 50)
    print("Demo completed! Check the frontend for real-time updates.")

if __name__ == "__main__":
    asyncio.run(main())