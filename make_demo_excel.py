import pandas as pd


def main() -> None:
    sales = pd.DataFrame(
        {
            "date": pd.to_datetime(
                ["2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08", "2026-01-09"]
            ),
            "region": ["North", "South", "North", "Central", "South"],
            "product": ["A", "B", "C", "A", "B"],
            "quantity": [10, 5, 12, 7, 9],
            "unit_price": [120000, 95000, 150000, 120000, 95000],
        }
    )
    sales["revenue"] = sales["quantity"] * sales["unit_price"]

    students = pd.DataFrame(
        {
            "student_id": ["S001", "S002", "S003", "S004", "S005"],
            "name": ["An", "Bình", "Chi", "Dũng", "Hà"],
            "class": ["10A1", "10A1", "10A2", "10A2", "10A1"],
            "math": [8.5, 7.0, 9.0, 6.5, 8.0],
            "english": [7.5, 8.0, 8.5, 6.0, 7.0],
        }
    )
    students["avg"] = (students["math"] + students["english"]) / 2

    path = "demo.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sales.to_excel(writer, index=False, sheet_name="Sales")
        students.to_excel(writer, index=False, sheet_name="Students")

    print(f"Created {path}")


if __name__ == "__main__":
    main()

