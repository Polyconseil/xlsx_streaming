.PHONY: test docs

update:
	pip install -r requirements_dev.txt

clean:
	find xlsx_streaming "(" -name '*.egg' -or -name '*.pyc' -or -name '*.pyo' ")" -delete
	find xlsx_streaming -type d "(" -name build -or -name __pycache__ -or -name _build ")" -exec rm -r {} \;

docs:
	sphinx-build -v -a -b html -d docs/_build/doctrees/ docs/ docs/_build/xlsx_streamingdoc/

test:
	pytest -v -Wdefault::DeprecationWarning

quality:
	python setup.py check --strict --metadata --restructuredtext
	pylint --reports=no setup.py xlsx_streaming tests
