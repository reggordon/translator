#!/bin/zsh
# Setup script for translator project

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install dependencies
echo "Installing dependencies from requirements.txt..."
pip install -r requirements.txt

echo "Setup complete!"
echo "Running test_translate.py using the virtual environment..."
source venv/bin/activate
python test_translate.py
echo "If you want to run other scripts, keep this terminal open or activate the environment again with: source venv/bin/activate"
