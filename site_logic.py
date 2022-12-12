from flask import Flask, render_template, send_file, request, send_from_directory
import os
import tracker_logic as tl

app = Flask(__name__, template_folder='templates', static_folder='static')

@app.route("/")
def home():
    return render_template('import_page.html')

@app.route('/index', methods = ['POST', 'GET'])
def upload_file():
    global results
    if request.method == 'POST':
        f = request.files['file']
        f.save(f.filename)
        file_path = os.path.abspath(f.filename)

        imported = tl.import_tmrs(file_path)
        current_tmr = tl.remove_old(imported)
        results = tl.import_movements(current_tmr)

        return render_template('index.html', data=results)

@app.route('/edit', methods = ['POST', 'GET'])
def edit_tmr():
    if request.method == 'POST':
        ind = int(request.form.get('index_button'))
    
    tmr_info = {
        'name': results[ind]['tmr name'],
        'num': results[ind]['tmr num'],
        'require': results[ind]['support needed'],
        'start': results[ind]['start dtg'],
        'pickup': results[ind]['pickup location'],
        'end': results[ind]['end dtg'],
        'dropoff': results[ind]['dropoff location'],
        'comments': results[ind]['comments'],
        'support': results[ind]['support unit'],
        'status': results[ind]['status']
    }
        
    return render_template('edit_page.html', data=tmr_info)

@app.route('/submit_edit', methods = ['POST'])
def submit_edit():
    if request.method == 'POST':
        name = request.form.get('name')
        num = request.form.get('num')
        req = request.form.get('require')
        start = request.form.get('start')
        pickup = request.form.get('pickup')
        end = request.form.get('end')
        dropoff = request.form.get('dropoff')
        comments = request.form.get('comments')
        support = request.form.get('support')
        status = request.form.get('status')

        for k, v in results.items():
            if num in v['tmr num']:
                results[k] = {
                    'tmr name': name,
                    'tmr num': num,
                    'support needed': req,
                    'start dtg': start,
                    'pickup location': pickup,
                    'end dtg': end,
                    'dropoff location': dropoff,
                    'support unit': support,
                    'status': status,
                    'comments': comments
                }

        tl.export_movement_tracker(results)
        return render_template('index.html', data=results)

@app.route('/download')
def export_tmr_file():
    tl.export_movement_tracker(results)
    file_excel = tl.format_excel()

    return send_file(file_excel, as_attachment=True)

if __name__ == "__main__":
    app.run()