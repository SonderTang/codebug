from flask import request, jsonify
from data_processing.bug_query import query_bug_data
# from data_processing.code_query import query_code_date, query_code_bug_date
# from utils.excel_utils import file_thousand_line_code_bug_rate

def register_routes(app):
    @app.route('/thousand_line_code_bug_rate', methods=['POST'])
    def thousand_line_code_bug_rate():
        year = request.values.get('year')
        month = request.values.get('month')
        query_bug_data()
        # query_code_date(year, month)
        # query_code_bug_date(year, month)
        # filename = file_thousand_line_code_bug_rate(year, month)
        return jsonify({'filename': filename})