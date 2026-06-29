.PHONY: setup
setup:
	@# A Python version is not explicitly specified via the --python flag because it is picked up from
	@# the special .python-version file.
	uv sync --dev

.PHONY: clean
clean:
	rm -rf build dist src/*.egg-info .ruff_cache
	find . -name '__pycache__' -exec rm -rf {} +
