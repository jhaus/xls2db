publish:
	python setup.py check
	python setup.py sdist
	python setup.py register sdist upload

test:
	python ./test.py
