install:
	@mkdir -p ~/.local/bin
	@ln -sf $(PWD)/clarifier.py ~/.local/bin/clarifier
	@echo "Installed Clarifier to ~/.local/bin/clarifier"
