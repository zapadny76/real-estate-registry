# Real Estate Registry

A desktop application for managing real estate records using PyQt6 and MongoDB.

## Features

- Management of houses, premises, and owners
- Import data from Excel files
- Track ownership history
- MongoDB database storage
- User-friendly interface

## Requirements

- Python 3.8+
- MongoDB
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/real-estate-registry.git
cd real-estate-registry
```

2. Create and activate virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Make sure MongoDB is running on localhost:27017

## Usage

Run the application:
```bash
python main.py
```

## Project Structure

```
real-estate-registry/
├── main.py              # Main application file
├── requirements.txt     # Python dependencies
├── README.md           # Project documentation
└── .gitignore          # Git ignore file
```

## License

MIT License
