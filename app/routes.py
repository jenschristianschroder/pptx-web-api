import os
from dotenv import load_dotenv
from flask import Blueprint, request, jsonify
from app.services.generate_pptx import generate_ppt, fetch_data

load_dotenv()
main = Blueprint('main', __name__)

@main.route('/generate-ppt', methods=['POST'])
def generate_presentation():
    data = request.json
    jobid = data.get('jobid')
    
    if not jobid:
        return jsonify({"error": "jobid is required"}), 400

    DATAVERSE_ENTITY = os.getenv('DATAVERSE_ENTITY')
    DATAVERSE_ENTITY_COLUMNS = os.getenv('DATAVERSE_ENTITY_COLUMNS')
    DATAVERSE_ENTITY_FILTER_COLUMN = os.getenv('DATAVERSE_ENTITY_FILTER_COLUMN')


    data = fetch_data(
        entity=DATAVERSE_ENTITY,
        select=[DATAVERSE_ENTITY_COLUMNS],
        filter_expr="".join([DATAVERSE_ENTITY_FILTER_COLUMN, f" eq '{jobid}'"])
    )

    try:
        records = data  # Implement this function to fetch data
        if not records:
            return jsonify({"error": "No records found for the given jobid"}), 404

        output_filename = f"{jobid}_report.pptx"
        generate_ppt(jobid=jobid, records=records, output_filename=output_filename)

        return jsonify({"message": "Presentation generated successfully", "filename": output_filename}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500