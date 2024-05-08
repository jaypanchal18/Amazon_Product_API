from flask import Flask, request, render_template, Response
from autoscraper import AutoScraper
import io
import xlsxwriter

app = Flask(__name__)
amazon_scraper = AutoScraper()
amazon_scraper.load('amazon_in.json')

def get_amazon_result(search_query):
    url = 'https://www.amazon.in/s?k=%s' % search_query
    result = amazon_scraper.get_result_similar(url, group_by_alias=True)
    return _aggregate_result(result)

def _aggregate_result(result):
    final_result = []
    for i in range(len(list(result.values())[0])):
        try:
            final_result.append({alias: result[alias][i] for alias in result})
        except:
            pass
    return final_result

@app.route('/', methods=['GET'])
def index():
    query = request.args.get('q', '')
    results = get_amazon_result(query)
    return render_template('index.html', query=query, results=results)

@app.route('/download', methods=['GET'])
def download():
    query = request.args.get('q', '')
    results = get_amazon_result(query)

    # Create an Excel workbook and add a worksheet
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    # Write header row
    headers = ['ImageUrl', 'Title', 'Price', 'Reviews','Ratings','MRP', 'Previous_Bought','About']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data rows
    for row, result in enumerate(results, start=1):
        for col, key in enumerate(headers):
            worksheet.write(row, col, result.get(key, ''))

    # Close the workbook
    workbook.close()

    # Prepare the response with the Excel file
    output.seek(0)
    response = Response(output.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Content-Disposition"] = "attachment; filename=amazon_search_results.xlsx"

    return response

if __name__ == '__main__':
    app.run(port=8080, host='0.0.0.0')
