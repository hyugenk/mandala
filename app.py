from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
import re


app = Flask(__name__, static_folder='assets', static_url_path='/assets')

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

file_output_global = None


# =========================
# FUNGSI BACA FILE (FIX XLSX ERROR)
# =========================
def baca_file_excel(file_path):

    if file_path.endswith(".xlsx"):
        return pd.read_excel(file_path, engine="openpyxl")

    elif file_path.endswith(".xls"):
        return pd.read_excel(file_path, engine="xlrd")

    elif file_path.endswith(".csv"):
        return pd.read_csv(file_path)

    else:
        raise Exception("Format file tidak didukung")


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/mergerdatapenjamin")
def mergerdatapenjamin():
    return render_template("mergerdatapenjamin.html")

@app.route("/process", methods=["POST"])
def process():

    global file_output_global

    try:
        file_utama = request.files["file_utama"]
        file_penjamin = request.files["file_penjamin"]

        path1 = os.path.join(UPLOAD_FOLDER, file_utama.filename)
        path2 = os.path.join(UPLOAD_FOLDER, file_penjamin.filename)

        file_utama.save(path1)
        file_penjamin.save(path2)

        # =========================
        # BACA FILE
        # =========================
        df_utama = baca_file_excel(path1)
        df_hp = baca_file_excel(path2)

        # =========================
        # NORMALISASI KOLOM
        # =========================
        df_utama.columns = df_utama.columns.str.strip()
        df_hp.columns = df_hp.columns.str.strip()

        # =========================
        # VALIDASI
        # =========================
        if 'Nama Penjamin' not in df_utama.columns:
            return jsonify({"status": "error", "message": "Kolom 'Nama Penjamin' tidak ada di file utama"})

        if 'Nama Penjamin' not in df_hp.columns:
            return jsonify({"status": "error", "message": "Kolom 'Nama Penjamin' tidak ada di file penjamin"})

        if 'NO HP' not in df_hp.columns:
            return jsonify({"status": "error", "message": "Kolom 'NO HP' tidak ada di file penjamin"})

        # =========================
        # NORMALISASI VALUE
        # =========================
        df_utama['KEY'] = df_utama['Nama Penjamin'].astype(str).str.strip().str.lower()
        df_hp['KEY'] = df_hp['Nama Penjamin'].astype(str).str.strip().str.lower()

        df_hp['NO HP'] = df_hp['NO HP'].astype(str).str.strip()

        # =========================
        # GROUP NO HP
        # =========================
        df_hp_grouped = (
            df_hp.groupby('KEY')['NO HP']
            .apply(lambda x: ' / '.join(sorted(set(x))))
            .reset_index()
        )

        # =========================
        # MERGE
        # =========================
        hasil_merge = df_utama.merge(df_hp_grouped, on='KEY', how='left')

        hasil_merge.drop(columns=['KEY'], inplace=True)

        # =========================
        # SIMPAN FILE
        # =========================
        output_file = os.path.join(
            UPLOAD_FOLDER,
            "hasil_merger_data_penjamin.xlsx"
        )

        file_output_global = output_file
        hasil_merge.to_excel(output_file, index=False)

        # =========================
        # PREVIEW
        # =========================
        table_html = hasil_merge.head(50).to_html(
            classes="table table-bordered table-striped",
            index=False
        )

        return jsonify({
            "status": "success",
            "table": table_html
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })
    



# =========================
# HALAMAN REKAP
# =========================
@app.route("/rekap")
def rekap():
    return render_template("rekap.html")


# =========================
# PROSES REKAP DATA REGISTER (POPUP)
# =========================
@app.route("/process_rekap", methods=["POST"])
def process_rekap():

    global file_output_global  # untuk download

    try:
        file = request.files["file_rekap"]

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        # =========================
        # BACA FILE (pakai fungsi aman)
        # =========================
        df = baca_file_excel(path)

        hasil = []

        for i in range(len(df)):

            no_register = df.loc[i, 'Unnamed: 3']

            if pd.notna(no_register):
                no_register_str = str(no_register).strip()

                # validasi format register
                if not re.search(r'\d', no_register_str):
                    continue

                if "NOMOR" in no_register_str.upper():
                    continue

                try:
                    nama_anak = df.loc[i, 'Unnamed: 6']
                    no_paspor = df.loc[i+2, 'Unnamed: 14']
                    masa_berlaku = df.loc[i+5, 'Unnamed: 14']
                    nama_penjamin = df.loc[i+7, 'Unnamed: 14']
                    alamat = df.loc[i+10, 'Unnamed: 14']

                    if str(nama_anak).strip().upper() == "NAMA":
                        continue

                    hasil.append({
                        "Nama Anak": nama_anak,
                        "No Register": no_register,
                        "Nama Penjamin": nama_penjamin,
                        "Alamat": alamat,
                        "No Paspor": no_paspor,
                        "Masa Berlaku": masa_berlaku
                    })

                except:
                    continue

        hasil_df = pd.DataFrame(hasil)

        # =========================
        # SIMPAN FILE
        # =========================
        output_file = os.path.join(
            UPLOAD_FOLDER,
            "hasil_rekap.xlsx"
        )

        file_output_global = output_file
        hasil_df.to_excel(output_file, index=False)

        # =========================
        # PREVIEW (50 DATA SAJA)
        # =========================
        table_html = hasil_df.head(50).to_html(
            classes="table table-bordered table-striped",
            index=False
        )

        return jsonify({
            "status": "success",
            "table": table_html
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })


@app.route("/download_file")
def download_file():
    global file_output_global
    return send_file(file_output_global, as_attachment=True)


# =========================
# REKAP KABUPATEN (MULTI SHEET + POPUP)
# =========================
from flask import jsonify  # 🔥 WAJIB ADA

@app.route("/process_kabupaten", methods=["POST"])
def process_kabupaten():
    global file_output_global

    try:
        file = request.files["file_kabupaten"]

        if not file.filename.endswith(('.xls', '.xlsx')):
            return jsonify({
                "status": "error",
                "message": "Gunakan file Excel (.xls/.xlsx)"
            })

        path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(path)

        df = baca_file_excel(path)
        df.columns = df.columns.str.strip().str.upper()

        # VALIDASI
        kolom_wajib = ["KOTA/KABUPATEN", "JENIS PERMOHONAN", "JENIS KELAMIN"]
        for k in kolom_wajib:
            if k not in df.columns:
                return jsonify({
                    "status": "error",
                    "message": f"Kolom '{k}' tidak ditemukan"
                })

        # NORMALISASI GENDER 🔥
        df["JENIS KELAMIN"] = df["JENIS KELAMIN"].astype(str).str.upper()

        df["JENIS KELAMIN"] = df["JENIS KELAMIN"].replace({
            "LAKI-LAKI": "L",
            "LAKI LAKI": "L",
            "PRIA": "L",
            "PEREMPUAN": "P",
            "WANITA": "P"
        })

        list_kota = df["KOTA/KABUPATEN"].dropna().unique()
        list_permohonan = df["JENIS PERMOHONAN"].dropna().unique()

        index_full = pd.MultiIndex.from_product(
            [list_kota, list_permohonan],
            names=["KOTA/KABUPATEN", "JENIS PERMOHONAN"]
        )

        pivot = pd.pivot_table(
            df,
            index=["KOTA/KABUPATEN", "JENIS PERMOHONAN"],
            columns="JENIS KELAMIN",
            aggfunc="size",
            fill_value=0
        )

        pivot = pivot.reindex(index_full, fill_value=0)
        pivot["JUMLAH"] = pivot.sum(axis=1)

        pivot = pivot.reset_index()
        pivot.insert(0, "No", range(1, len(pivot) + 1))
        pivot.columns.name = None

        # =========================
        # SIMPAN MULTI SHEET
        # =========================
        output_file = os.path.join(UPLOAD_FOLDER, "rekap_kabupaten.xlsx")
        file_output_global = output_file

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

            pivot.to_excel(writer, sheet_name="ALL_DATA", index=False)

            for kota in list_kota:
                data_kota = pivot[pivot["KOTA/KABUPATEN"] == kota].copy()
                data_kota["No"] = range(1, len(data_kota) + 1)
                sheet_name = str(kota)[:31]
                data_kota.to_excel(writer, sheet_name=sheet_name, index=False)

        # =========================
        # PREVIEW
        # =========================
        table_html = pivot.head(50).to_html(
            classes="table table-bordered table-striped",
            index=False
        )

        # =========================
        # GRAFIK
        # =========================
        grafik_data = (
            pivot.groupby("KOTA/KABUPATEN")
            .sum(numeric_only=True)
            .reset_index()
        )

        labels = grafik_data["KOTA/KABUPATEN"].tolist()
        laki = grafik_data.get("L", pd.Series([0]*len(labels))).tolist()
        perempuan = grafik_data.get("P", pd.Series([0]*len(labels))).tolist()
        total = grafik_data["JUMLAH"].tolist()

        return jsonify({
            "status": "success",
            "table": table_html,
            "chart": {
                "labels": labels,
                "laki": laki,
                "perempuan": perempuan,
                "total": total
            }
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })





@app.route("/download_kabupaten")
def download_kabupaten():
    global file_output_global
    return send_file(file_output_global, as_attachment=True)



@app.route("/rekap_kabupaten")
def rekap_kabupaten():
    return render_template("rekap_kabupaten.html")



# =========================
# Run aplikasi
# =========================
if __name__ == "__main__":
    app.run(debug=True)