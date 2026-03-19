import pytest

import excelipy as ep
from excelipy.service import unnest_components


def test_group_unnesting():
    comps = [
        ep.Group(components=[ep.Fill()]),
        ep.Group(components=[ep.Fill(), ep.Group(components=[ep.Fill()])]),
        ep.Group(
            components=[
                ep.Fill(),
                ep.Group(components=[ep.Group(components=[ep.Fill()])]),
            ]
        ),
    ]
    out = unnest_components(comps)
    size = len(out)
    for c in out:
        assert not isinstance(c, ep.Group)
    assert size == 5


if __name__ == "__main__":
    pytest.main([__file__])
