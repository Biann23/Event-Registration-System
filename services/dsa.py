from collections import deque

def merge_sort(items, key=lambda x: x):
    """Custom merge sort (O(n log n))."""
    items = list(items)
    if len(items) < 2:
        return items
    mid = len(items) // 2
    left, right = merge_sort(items[:mid], key), merge_sort(items[mid:], key)
    out, i, j = [], 0, 0
    while i < len(left) and j < len(right):
        if key(left[i]) <= key(right[j]):
            out.append(left[i]); i += 1
        else:
            out.append(right[j]); j += 1
    return out + left[i:] + right[j:]

def binary_search_title(sorted_events, title):
    """Exact-title binary search; input must be sorted by title.casefold()."""
    target, lo, hi = title.casefold(), 0, len(sorted_events) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        value = sorted_events[mid].title.casefold()
        if value == target: return sorted_events[mid]
        if value < target: lo = mid + 1
        else: hi = mid - 1
    return None

class RegistrationHashIndex:
    """Dictionary-backed hash lookup keyed by (student_id, event_id)."""
    def __init__(self, registrations=()):
        self.data = {(r.student_id, r.event_id): r for r in registrations}
    def get(self, student_id, event_id):
        return self.data.get((student_id, event_id))
    def contains(self, student_id, event_id):
        return (student_id, event_id) in self.data

class WaitingQueue:
    """FIFO queue used to demonstrate queue operations."""
    def __init__(self): self.items = deque()
    def enqueue(self, value): self.items.append(value)
    def dequeue(self): return self.items.popleft() if self.items else None
    def __len__(self): return len(self.items)
