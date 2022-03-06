
install_deps:
	pip3 install -r requirements.txt

freeze:
	pip3 freeze > requirements.txt

run:
	python3 main.py

revenue:
	python3 revenue.py

account_receivable:
	python3 account_receivable.py