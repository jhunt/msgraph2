default: build

build:
	python3 -m build --wheel

clean:
	rm -rf dist/*

testpub:
	twine upload --repository testpypi dist/*

pub:
	twine upload dist/*

.PHONY: build clean testpub pub
