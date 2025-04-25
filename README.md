# pptx-web-api

## Overview
This project is a Flask application that generates PowerPoint presentations from Microsoft Dataverse data using a fixed template. It retrieves data from Dataverse, processes it, and populates a PowerPoint file accordingly.

## Project Structure
```
pptx-web-api
├── app
│   ├── __init__.py
│   ├── routes.py
│   ├── services
│   │   └── generate_pptx.py
│   └── utils
│       └── __init__.py
├── wsgi.py
├── requirements.txt
├── .env.example
└── README.md
```

## Setup Instructions

1. **Clone the Repository**
   ```bash
   git clone <repository-url>
   cd pptx-web-api
   ```

2. **Create a Virtual Environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure Environment Variables**
   Copy the `.env.example` file to `.env` and fill in the required values:
   ```
   DATAVERSE_CLIENT_ID=<your-client-id>
   DATAVERSE_CLIENT_SECRET=<your-client-secret>
   DATAVERSE_TENANT_ID=<your-tenant-id>
   DATAVERSE_URL=<your-dataverse-url>
   ```

## Usage

1. **Run the Application**
   ```bash
   python wsgi.py
   ```

2. **Access the API**
   The application will be available at `http://localhost:5000`. You can define routes in `app/routes.py` to handle requests for generating PowerPoint presentations.

## Contributing
Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.