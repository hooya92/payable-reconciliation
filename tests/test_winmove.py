"""Window move / resize arithmetic (the Windows calls around it are thin)."""
import unittest

from webui.winmove import geometry

RECT = (100, 50, 500, 800)
MIN = (460, 640)


class GeometryTests(unittest.TestCase):
    def test_move_follows_the_cursor_and_keeps_the_size(self):
        self.assertEqual(geometry("move", RECT, (300, 60), (350, 90), MIN), (150, 80, 500, 800))

    def test_corner_resize_grows_and_shrinks_both_ways(self):
        self.assertEqual(geometry("se", RECT, (600, 850), (640, 900), MIN), (100, 50, 540, 850))
        self.assertEqual(geometry("se", RECT, (600, 850), (580, 840), MIN), (100, 50, 480, 790))

    def test_edges_change_only_one_dimension(self):
        self.assertEqual(geometry("s", RECT, (0, 0), (30, 40), MIN), (100, 50, 500, 840))
        self.assertEqual(geometry("e", RECT, (0, 0), (30, 40), MIN), (100, 50, 530, 800))

    def test_the_window_never_gets_smaller_than_its_minimum(self):
        self.assertEqual(geometry("se", RECT, (0, 0), (-900, -900), MIN), (100, 50, 460, 640))


if __name__ == "__main__":
    unittest.main()
