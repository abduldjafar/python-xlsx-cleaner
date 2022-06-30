from pandas import read_excel, concat, ExcelWriter
from dateutil import parser
import argparse


if __name__ == "__main__":

    my_parser = argparse.ArgumentParser(allow_abbrev=False)
    my_parser.add_argument("--xlsx_file", action="store", type=str, required=True)
    my_parser.add_argument("--sheet_name", action="store", type=str, required=True)
    my_parser.add_argument(
        "--xlsx_destination", action="store", type=str, required=True
    )

    args = my_parser.parse_args()

    xlsx_file = args.xlsx_file
    sheet_name = args.sheet_name
    xlsx_destination = args.xlsx_destination

    data = read_excel(xlsx_file, sheet_name=sheet_name)
    all_datas_cleaned = []

    for column in ["Category", "Amount", "Date"]:
        if column == "Category":
            data_cleaned = data[["Order ID", column]]
        else:
            data_cleaned = data[[column]]

        data_cleaned[column] = data_cleaned[column].apply(lambda x: x.split("|"))
        data_cleaned = data_cleaned.explode(column)

        if column == "Date":
            data_cleaned[column] = data_cleaned[column].apply(
                lambda x: str(parser.parse(x).date())
            )
        else:
            data_cleaned[column] = data_cleaned[column].apply(lambda x: x.strip())

        data_cleaned = data_cleaned.reset_index(drop=True)

        all_datas_cleaned.append(data_cleaned)

    df_concat = concat(all_datas_cleaned, axis=1)

    writer = ExcelWriter(xlsx_destination, engine="xlsxwriter")

    df_concat.to_excel(writer, sheet_name="Sheet1", index=False)

    writer.save()
