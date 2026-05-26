import pandas as pd
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

    for c in out:
        assert isinstance(c.name, str)

    assert ep.Group().name == "group"
    assert ep.Table(data=pd.DataFrame()).name == "table"
    assert ep.Group(name="group_name").name == "group_name"

    assert size == 5


if __name__ == "__main__":
    pytest.main([__file__])
