
all: start

start: clean summary revenue account_receivable
	zip -rm dist dist

clean:
	rm -rf dist
	mkdir dist

summary:
	python3 main.py

revenue:
	python3 revenue.py

account_receivable:
	python3 account_receivable.py


install_deps:
	pip3 install -r requirements.txt

freeze:
	pip3 freeze > requirements.txt

step_up:
	python3 -m venv .env

active:
	source .env/bin/activate
	
.PHONY: summary clean revenue account_receivable