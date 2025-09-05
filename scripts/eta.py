"""
ETA (Estimated Time of Arrival) calculation utilities.
"""
import time

class ETAEstimator:
    """Exponential moving average estimator for time per item."""
    
    def __init__(self, alpha=0.25):
        self.alpha = alpha
        self.sec_per_item = None
    
    def update(self, time_per_item):
        """Update the moving average with new timing data."""
        if self.sec_per_item is None:
            self.sec_per_item = time_per_item
        else:
            self.sec_per_item = self.alpha * time_per_item + (1 - self.alpha) * self.sec_per_item

def fmt_hms(seconds):
    """Format seconds as H:MM:SS."""
    if seconds is None or seconds < 0:
        return "??:??:??"
    
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    
    if hours > 0:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    else:
        return f"{minutes}:{secs:02d}"