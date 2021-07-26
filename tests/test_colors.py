import pytest

from gslides.colors import Palette


class TestPalette:
    @pytest.mark.xfail(reason=ValueError)
    def test_init_fail(self):
        Palette("google")
        assert True

    def test_init_success(self):
        Palette("prodigy")
        assert True
