import pandas as pd


def check_count():
	target_file = "data/marcap-2021.csv.gz"
	df = pd.read_csv(target_file)

	date_check = df.groupby("Date").count()
	print(date_check)


if __name__ == "__main__":
	check_count()