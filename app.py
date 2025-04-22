from flask import Flask, render_template, redirect, url_for
import mysql.connector
from form_laporan import laporan_routes
from form_perjadin import perjadin_routes
from form_laporan import generate
from flask import Flask, send_file, Blueprint, render_template, request, redirect, url_for

app = Flask(__name__)


app.register_blueprint(laporan_routes)
app.register_blueprint(perjadin_routes)

# Fungsi helper untuk koneksi database (bukan sebagai route!)
def get_connection():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="db_perjadin"
    )

@app.route('/')
def home():
    # Atau bisa return halaman home
    return redirect(url_for('dashboard'))  # redirect ke dashboard

@app.route('/dashboard')
def dashboard():
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM history_file ORDER BY tanggal_generate DESC")
    history = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('dashboard.html', history=history)

@app.route('/dashboard/form_lapangan')
def form_lapangan():
    return render_template('form_lapangan.html')

@app.route('/dashboard/form_perjadin')
def form_perjadin():
    return render_template('form_perjadin.html')
@app.route('/generate', methods=['POST'])
def generate():
    try:
        # Path ke file Word yang dihasilkan oleh generate_laporan
        file_path = generate

        # Kirim file ke browser sebagai attachment (download)
        return send_file(file_path, as_attachment=True)

    except Exception as e:
        return f"Gagal mengunduh laporan: {str(e)}"




if __name__ == '__main__':
    app.run(debug=True)
