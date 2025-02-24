from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import numpy as np
import numpy_financial as npf
import matplotlib.pyplot as plt
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULTS_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)


def calcular_flujo_caja(df, absorcion, porcentaje_reserva, porcentaje_desembolso, meses_reserva, meses_lag_desembolso):
    distribucion = []
    mes_actual = 1
    for i, row in df.iterrows():
        if i % absorcion == 0 and i != 0:
            mes_actual += 1
        distribucion.append((row["unidad"], row["precio"], mes_actual))

    flujo_caja = []
    max_mes = 0
    for unidad, precio, mes_venta in distribucion:
        reserva = precio * (porcentaje_reserva / 100)
        desembolso = precio * (porcentaje_desembolso / 100)
        cuota_reserva_mensual = reserva / meses_reserva

        pagos_reserva = {f"mes {m}": 0 for m in range(1, mes_venta + meses_reserva + meses_lag_desembolso + 1)}
        for m in range(mes_venta, mes_venta + meses_reserva):
            pagos_reserva[f"mes {m}"] = cuota_reserva_mensual

        mes_desembolso = mes_venta + meses_reserva + meses_lag_desembolso
        pagos_reserva[f"mes {mes_desembolso}"] = desembolso

        flujo_caja.append({
            "unidad": unidad,
            "precio": precio,
            "Reserva": reserva,
            "Desembolso": desembolso,
            "Cuota Reserva Mensual": cuota_reserva_mensual,
            **pagos_reserva
        })

        max_mes = max(max_mes, mes_desembolso)

    df_flujo_caja = pd.DataFrame(flujo_caja)

    # Agregar sumatoria de flujos mensuales
    sumatorias = {col: df_flujo_caja[col].replace('[\$,]', '', regex=True).astype(float).sum()
                  if "mes" in col else "" for col in df_flujo_caja.columns}
    sumatorias["unidad"] = "Total Flujo Mensual"

    df_sumatoria = pd.DataFrame([sumatorias])
    df_flujo_caja = pd.concat([df_flujo_caja, df_sumatoria], ignore_index=True)

    return df_flujo_caja


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Obtener datos del formulario
        absorcion = int(request.form["absorcion"])
        porcentaje_reserva = float(request.form["porcentaje_reserva"])
        porcentaje_desembolso = float(request.form["porcentaje_desembolso"])
        meses_reserva = int(request.form["meses_reserva"])
        meses_lag_desembolso = int(request.form["meses_lag_desembolso"])

        # Guardar archivo cargado
        file = request.files["file"]
        if file.filename == "":
            return redirect(request.url)

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Leer archivo y calcular flujo de caja
        df = pd.read_excel(filepath)
        df_flujo_caja = calcular_flujo_caja(df, absorcion, porcentaje_reserva,
                                             porcentaje_desembolso, meses_reserva, meses_lag_desembolso)

        # Guardar archivo de salida
        output_file = os.path.join(RESULTS_FOLDER, "flujo_caja.xlsx")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_flujo_caja.to_excel(writer, index=False, sheet_name="Flujo de Caja")
            workbook = writer.book
            worksheet = writer.sheets["Flujo de Caja"]
            num_format = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column("B:Z", 15, num_format)

        # Generar gráfico
        suma_flujos = df_flujo_caja.iloc[:, 5:].sum()
        fig, ax = plt.subplots(figsize=(10, 5))
        suma_flujos.plot(kind='bar', ax=ax, color='skyblue')
        ax.set_title("Distribución de Flujos de Caja por Mes")
        ax.set_xlabel("Mes")
        ax.set_ylabel("Total Recibido")
        ax.grid(axis="y", linestyle="--", alpha=0.7)

        # Guardar el gráfico
        graph_path = os.path.join(RESULTS_FOLDER, "flujo_caja.png")
        fig.savefig(graph_path, dpi=300, bbox_inches="tight")

        return render_template("result.html", graph_path=graph_path)

    return render_template("index.html")


@app.route("/download")
def download():
    return send_file(os.path.join(RESULTS_FOLDER, "flujo_caja.xlsx"), as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
