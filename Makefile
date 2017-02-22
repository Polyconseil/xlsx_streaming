.PHONY: test docs

update:
	pip install -r requirements_dev.txt

clean:
	find xlsx_streaming "(" -name '*.egg' -or -name '*.pyc' -or -name '*.pyo' ")" -delete
	find xlsx_streaming -type d "(" -name build -or -name __pycache__ -or -name _build ")" -exec rm -r {} \;

docs:
	python setup.py build_sphinx

test:
	python setup.py test
