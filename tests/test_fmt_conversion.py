import pytest

from excelipy.const import python_to_excel_fmt


@pytest.mark.parametrize(
    "python_fmt, expected_excel",
    [
        # --- Numeric Formats ---
        ("", "General"),
        (",.0f", "#,##0"),
        (",.2f", "#,##0.00"),
        (".0f", "0"),
        (".2f", "0.00"),
        ("General", "General"),
        ("d", "0"),
        # --- Percentages ---
        ("%", "0%"),
        (".0%", "0%"),
        (".1%", "0.0%"),
        # --- Date Formats ---
        ("%H:%M:%S", "hh:mm:ss"),
        ("%Y-%m-%d", "yyyy-mm-dd"),
        ("%b %d, %Y", "mmm dd, yyyy"),
        ("%d/%m/%y", "dd/mm/yy"),
        # --- Excel formats ---
        ("#,##0", "#,##0"),
        ("#,##0.00", "#,##0.00"),
        ("0", "0"),
        ("0%", "0%"),
        ("0.0%", "0.0%"),
        ("0.00", "0.00"),
        ("General", "General"),
        ("dd/mm/yy", "dd/mm/yy"),
        ("hh:mm:ss", "hh:mm:ss"),
        ("mmm dd, yyyy", "mmm dd, yyyy"),
        ("yyyy-mm-dd", "yyyy-mm-dd"),
    ],
)
def test_python_to_excel_fmt(python_fmt, expected_excel):
    assert python_to_excel_fmt(python_fmt) == expected_excel


if __name__ == "__main__":
    pytest.main([__file__])
