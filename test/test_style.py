import pytest

import excelipy as ep


def test_style():
    temp_style = ep.Style()
    assert temp_style.pr() == 0
    assert temp_style.pl() == 0
    assert temp_style.pb() == 0
    assert temp_style.pt() == 0


if __name__ == "__main__":
    pytest.main([__file__])
