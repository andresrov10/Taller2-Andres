# Importar librerias
from matplotlib.patches import Patch
import pulp as lp
import pandas as pd
import matplotlib.pyplot as plt
from tabulate import tabulate

# Archivo de datos
file_name = 'data.xlsx'

# --------------------------------
# Datos
# --------------------------------
# Tablas
tabla2 = pd.read_excel(io=file_name, sheet_name='Tabla 2', header=0, index_col=0)
tabla3 = pd.read_excel(io=file_name, sheet_name='Tabla 3', header=0, index_col=0)

# --------------------------------
# Conjuntos
# --------------------------------
# Partes
I = list(tabla2.columns)

# Maquinas
J = list(tabla3.index)

# --------------------------------
# Parámetros
# --------------------------------

# Demanda de cada tipo de parte
d = {"Motores": 3503,
     "Airbags": 3008,
     "Ejes": 4003,
     "Amortiguadores": 4001,
     "Frenos": 2808
    }

# Factor de rendimiento para cada tipo de parte
f = {i: tabla2.loc["Rendimiento", i] for i in I}

# Tasa de producción de la parte i en la máquina j
k = {(i, j): float(tabla2.loc[j, i]) if not pd.isna(tabla2.loc[j, i]) else 0 for i in I for j in J }

# Tiempo de alistamiento de la máquina j para producir la parte i
t = {(i, j): float(tabla3.loc[j, i]) if not pd.isna(tabla3.loc[j, i]) else 0 for i in I for j in J }

# Numero de horas regulares para producir cualquier parte en cualquier maquina
# 8 horas por turno, 3 turnos por día, 5 días por semana
r = 8*3*5

# Numero de horas extra para producir cualquier parte en cualquier maquina
# 8 horas por turno, 3 turnos por día, 2 días por semana
e = 8*3*2 

# Partes producidas en cada máquina
solucion_actual = {(i, j): 0 for i in I for j in J}

solucion_actual[("Motores", 1)] = 2530
solucion_actual[("Motores", 2)] = 3309
solucion_actual[("Airbags", 2)] = 1325
solucion_actual[("Airbags", 3)] = 3022
solucion_actual[("Airbags", 4)] = 1123
solucion_actual[("Ejes", 4)] = 5338
solucion_actual[("Amortiguadores", 1)] = 4085
solucion_actual[("Amortiguadores", 5)] = 2071
solucion_actual[("Frenos", 3)] = 200
solucion_actual[("Frenos", 5)] = 4480

# --------------------------------
# Inferir Variables de decisión
# --------------------------------

# Si se produce la parte i en la máquina j, entonces 1, de lo contrario 0
z = {(i, j): 1 if solucion_actual[(i, j)] > 0 else 0 for i in I for j in J}

# Revisar si se cumple la demanda de cada parte
cumple_demanda = {i: sum(solucion_actual[(i, j)] for j in J) >= d[i] for i in I}

print("Se cumple la demanda de cada parte")
print(tabulate(cumple_demanda.items(), headers=["Parte", "Cumple"], tablefmt="fancy_grid"))

# Horas totales de producción de cada máquina
h = {j: sum(solucion_actual[(i, j)]/k[(i, j)] if k[i, j] > 0 else 0 for i in I) for j in J}

print("Horas totales de producción de cada máquina")
print(tabulate(h.items(), headers=["Máquina", "Horas"], tablefmt="fancy_grid"))

# Horas totales de producción de cada parte en cada máquina
h_ij = {(i, j): solucion_actual[(i, j)]/k[(i, j)] if k[i, j] > 0 else 0 for i in I for j in J}

for j in J:
    for i in I:
        if k[i, j] > 0:
            print("Horas totales de producción de la parte {} en la máquina {}".format(i, j))
            print(h_ij[(i, j)])

# Horas regulares disponibles para cada máquina
r_j = {j: r - sum(t[i, j]*z[i,j] for i in I) for j in J}

print(r_j)

# Repartir las horas de producción en horas regulares y horas extra para cada máquina y parte
x = {(i, j): 0 for i in I for j in J}
y = {(i, j): 0 for i in I for j in J}

x[("Amortiguadores", 1)] = h_ij[("Amortiguadores", 1)]
x[("Motores", 1)] = r_j[1]-x[("Amortiguadores", 1)] if r_j[1]-x[("Amortiguadores", 1)] <= h_ij[("Motores", 1)] else h_ij[("Motores", 1)]
y[("Motores", 1)] = h_ij[("Motores", 1)]-x[("Motores", 1)]

x[("Motores", 2)] = h_ij[("Motores", 2)]
x[("Airbags", 2)] = r_j[2]-x[("Motores", 2)] if r_j[2]-x[("Motores", 2)] <= h_ij[("Airbags", 2)] else h_ij[("Airbags", 2)]
y[("Airbags", 2)] = h_ij[("Airbags", 2)]-x[("Airbags", 2)]

x[("Airbags", 3)] = r_j[3]
y[("Airbags", 3)] = h_ij[("Airbags", 3)]-x[("Airbags", 3)]
y[("Frenos", 3)] = h_ij[("Frenos", 3)]

x[("Ejes", 4)] = r_j[4]
y[("Ejes", 4)] = h_ij[("Ejes", 4)]-x[("Ejes", 4)]
y[("Amortiguadores", 4)] = h_ij[("Airbags", 4)]

x[("Frenos", 5)] = h_ij[("Frenos", 5)]
x[("Amortiguadores", 5)] = r_j[5]-x[("Frenos", 5)] if r_j[5]-x[("Frenos", 5)] <= h_ij[("Amortiguadores", 5)] else h_ij[("Amortiguadores", 5)]
y[("Amortiguadores", 5)] = h_ij[("Amortiguadores", 5)]-x[("Amortiguadores", 5)]

print(f"Suma de horas extra: {sum(y[i, j] for i in I for j in J)}")

# -----------------
# Gráficos
# -----------------

# Generar un grafico de barras con las horas regulares y horas extra de producción de cada máquina
plt.figure(figsize=(10, 5))
plt.bar([j for j in J], [sum(x[i, j] for i in I) for j in J], label="Horas regulares", color="forestgreen")
plt.bar([j for j in J], [sum(y[i, j] for i in I) for j in J], bottom=[sum(x[i, j] for i in I) for j in J], label="Horas extra", color="sandybrown")
plt.ylim(0, 160)
plt.xlabel("Maquina")
plt.ylabel("Horas")
plt.title("Horas regulares y horas extra de producción por máquina")
# mostrar el valor de cada barra
for j in J:
    plt.text(j, sum(x[i, j] for i in I)/2, str(round(sum(x[i, j] for i in I), 2)))
    plt.text(j, sum(x[i, j] for i in I) + sum(y[i, j] for i in I)/2, str(round(sum(y[i, j] for i in I), 2)))
plt.legend()
plt.show()

# Generar un unico grafico de barras que muestre las horas regulares y horas extra de producción de cada parte en cada máquina
plt.figure(figsize=(12, 5))
bottom_r_limits = [0 for j in J]
bottom_e_limits = [sum(x[i, j] for i in I) for j in J]
colores_fuertes = {
    "Motores": "firebrick",
    "Airbags": "forestgreen",
    "Ejes": "royalblue",
    "Amortiguadores": "gold",
    "Frenos": "darkorange"
}

colores_claros = {
    "Motores": "lightpink",
    "Airbags": "limegreen",
    "Ejes": "skyblue",
    "Amortiguadores": "khaki",
    "Frenos": "lightsalmon"
}
for i in I:
    plt.bar([j for j in J], [x[i, j] for j in J], bottom=bottom_r_limits, color=colores_fuertes[i])
    bottom_r_limits = [bottom_r_limits[j] + x[i, j+1] for j in range(len(J))]
    for j in J:
        # Graficar el valor de cada barra para cada parte y maquina
        plt.text(J[j-1]+0.1, bottom_r_limits[j-1]-5, round(x[i, j],2), color="black") if x[i, j] > 5 else None
        plt.text(J[j-1]+0.2, bottom_r_limits[j-1]+2, round(x[i, j],2), color="black") if x[i, j] > 0 and x[i,j] <= 5 else None



for i in I:
    plt.bar([j for j in J], [y[i, j] for j in J], bottom=bottom_e_limits, color=colores_claros[i])
    bottom_e_limits = [bottom_e_limits[j] + y[i, j+1] for j in range(len(J))]
    for j in J:
        # Graficar el valor de cada barra para cada parte y maquina
        plt.text(J[j-1]-0.4, bottom_e_limits[j-1]-5, round(y[i, j],2), color="black") if y[i, j] > 5 else None
        plt.text(J[j-1]-0.4, bottom_e_limits[j-1]+2, f"{round(y[i, j],2)}-{i}", color="black") if y[i, j] > 0 and y[i,j] <= 5 else None

# Graficar dos leyendas, una para las horas regulares y otra para las horas extra
leyenda = [ Patch(color=colores_fuertes[i], label=f"{i}-Horas Regulares") for i in I]
leyenda2 = [Patch(color=colores_claros[i], label=f"{i}-Horas Extra") for i in I]
leyendaFinal = leyenda + leyenda2
# Agregar espacio a la derecha del grafico para mostrar las leyendas
plt.subplots_adjust(right=0.7)
# Graficar la leyenda a la derecha del grafico, fuera del area del grafico
plt.legend(handles=leyendaFinal, bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)

plt.ylim(0, 160)

plt.xlabel("Maquina")
plt.ylabel("Horas")
plt.title("Horas regulares y horas extra de producción por parte y máquina")
plt.show()

# Generar un grafico de barras con las horas de alistamiento de cada máquina para cada parte
plt.figure(figsize=(40, 50))
bottom_limit = [0 for j in J]
partes_graficadas = []
for i in I:
    plt.bar([j for j in J], [t[i, j]*z[i, j] for j in J], bottom=bottom_limit, label=i, color=colores_claros[i])
    bottom_limit = [bottom_limit[j] + t[i, j+1]*z[i, j+1] for j in range(len(J))]
    for j in J:
        # Graficar el valor de cada barra para cada parte y maquina
        plt.text(J[j-1], bottom_limit[j-1]-5, round(t[i, j]), color="black") if z[i,j] > 0.5 else None

plt.ylim(0, 40)
print(partes_graficadas)
plt.xlabel("Maquina")
plt.ylabel("Horas")
plt.title("Horas de alistamiento por parte y máquina")
plt.legend()
plt.show()



