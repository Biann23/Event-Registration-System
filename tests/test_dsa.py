import unittest
from dataclasses import dataclass
from services.dsa import merge_sort, binary_search_title, RegistrationHashIndex, WaitingQueue

@dataclass
class EventStub:
    title: str

@dataclass
class RegistrationStub:
    student_id: int
    event_id: int

class DSATests(unittest.TestCase):
    def test_merge_sort_returns_sorted_copy(self):
        source = [EventStub("Workshop"), EventStub("assembly"), EventStub("Festival")]
        result = merge_sort(source, key=lambda event: event.title.casefold())
        self.assertEqual([item.title for item in result], ["assembly", "Festival", "Workshop"])
        self.assertEqual([item.title for item in source], ["Workshop", "assembly", "Festival"])

    def test_binary_search_finds_exact_title_case_insensitively(self):
        events = merge_sort([EventStub("Workshop"), EventStub("Assembly")], key=lambda event: event.title.casefold())
        self.assertEqual(binary_search_title(events, "assembly").title, "Assembly")
        self.assertIsNone(binary_search_title(events, "assembly party"))

    def test_registration_hash_index_looks_up_student_event_pair(self):
        registration = RegistrationStub(student_id=7, event_id=12)
        index = RegistrationHashIndex([registration])
        self.assertTrue(index.contains(7, 12))
        self.assertIs(index.get(7, 12), registration)
        self.assertFalse(index.contains(7, 13))

    def test_waiting_queue_is_fifo(self):
        queue = WaitingQueue()
        queue.enqueue("first")
        queue.enqueue("second")
        self.assertEqual(len(queue), 2)
        self.assertEqual(queue.dequeue(), "first")
        self.assertEqual(queue.dequeue(), "second")
        self.assertIsNone(queue.dequeue())

if __name__ == "__main__":
    unittest.main()
