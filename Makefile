install:
	python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

dry-run:
	python -m src.main --month 2025-08 --config config/config.yml --out output/merged --downloads data/downloads --dry-run

run:
	python -m src.main --month 2025-08 --config config/config.yml --out output/merged --downloads data/downloads
