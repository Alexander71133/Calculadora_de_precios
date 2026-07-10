import customtkinter as ctk
import pandas as pd
from datetime import datetime
import os

# Configuración de colores ALEXSolutions
AZUL_CORP = "#1B365D"
AMARILLO_CORP = "#FFD700"
BLANCO_FONDO = "#F8F9FA"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class CalculadoraALEXSolutions(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ALXSolutions - Calculo Ganancia")
        self.geometry("460x710") 
        self.resizable(False, False)
        self.configure(fg_color=BLANCO_FONDO)

        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(expand=True, fill="both", padx=30, pady=(10, 10))

        ctk.CTkLabel(self.main_frame, text="ALXSolutions", font=("Roboto", 20, "bold"), text_color=AZUL_CORP).pack(pady=(5, 0))

        self.entry_costo = ctk.CTkEntry(self.main_frame, placeholder_text="Costo del Producto ($)", height=40, font=("Roboto", 16), border_color=AZUL_CORP)
        self.entry_costo.pack(pady=5, fill="x")
        self.entry_costo.focus()

        self.entry_margen = ctk.CTkEntry(self.main_frame, placeholder_text="Margen Deseado (%)", height=40, font=("Roboto", 16), border_color=AZUL_CORP)
        self.entry_margen.pack(pady=5, fill="x")

        self.res_frame = ctk.CTkFrame(self.main_frame, fg_color=AZUL_CORP, corner_radius=10)
        self.res_frame.pack(pady=10, fill="x")

        self.label_precio = ctk.CTkLabel(self.res_frame, text="PRECIO VENTA: $0", font=("Roboto", 20, "bold"), text_color=AMARILLO_CORP)
        self.label_precio.pack(pady=(10, 2))

        self.label_ganancia = ctk.CTkLabel(self.res_frame, text="GANANCIA: $0", font=("Roboto", 18), text_color="white")
        self.label_ganancia.pack(pady=(2, 10))

        self.btn_calcular = ctk.CTkButton(self.main_frame, text="REGISTRAR CÁLCULO", command=self.procesar, height=45, font=("Roboto", 16, "bold"), fg_color=AMARILLO_CORP, text_color=AZUL_CORP, hover_color="#E6C200")
        self.btn_calcular.pack(pady=10, fill="x")

        self.bind('<Return>', lambda event: self.procesar())

        ctk.CTkLabel(self.main_frame, text="Historial Reciente (Últimos 10)", font=("Roboto", 14, "bold"), text_color=AZUL_CORP).pack(pady=(10, 5), anchor="w")
        
        self.tabla = ctk.CTkTextbox(self.main_frame, height=80, font=("Courier New", 11), border_width=1, border_color=AZUL_CORP)
        self.tabla.pack(fill="both", expand=True)
        
        self.btn_limpiar = ctk.CTkButton(self.main_frame, text="Limpiar Historial Excel", command=self.limpiar_historial, height=25, font=("Roboto", 11), fg_color="#D32F2F", text_color="white", hover_color="#B71C1C")
        self.btn_limpiar.pack(pady=(10, 5))
        
        self.actualizar_tabla()

    def formato_moneda(self, valor):
        return "{:,.0f}".format(valor).replace(",", ".")

    def actualizar_tabla(self):
        self.tabla.configure(state="normal")
        self.tabla.delete("1.0", "end")
        archivo = "registro_precios.xlsx"
        
        header = f"{'FECHA':<6} | {'COSTO':<11} | {'VENTA':<11} | {'GANANCIA':<10} | {'MG%'}\n"
        sep = "-" * 60 + "\n"
        self.tabla.insert("end", header + sep)

        if os.path.isfile(archivo):
            try:
                df = pd.read_excel(archivo)
                # Tomamos los últimos 10 pero mantenemos el orden de llegada (head no, tail sí)
                # Para que el primero del grupo de 10 quede arriba y los nuevos abajo
                ultimos = df.tail(10) 
                
                for _, fila in ultimos.iterrows():
                    fecha_str = str(fila.get('Fecha', ''))
                    fecha_formato = f"{fecha_str[8:10]}-{fecha_str[5:7]}"
                    
                    costo_f = self.formato_moneda(fila['Costo'])
                    venta_f = self.formato_moneda(fila['Precio Venta'])
                    ganancia_f = self.formato_moneda(fila['Ganancia'])
                    mg_f = f"{int(fila.get('Margen %', 0))}%"
                    
                    texto = f"{fecha_formato:<6} | {costo_f:<11} | {venta_f:<11} | {ganancia_f:<10} | {mg_f}\n"
                    self.tabla.insert("end", texto)
            except Exception as e:
                print(f"Error: {e}")
        
        self.tabla.configure(state="disabled")

    def procesar(self):
        try:
            costo_str = self.entry_costo.get().replace(".", "").replace(",", "")
            costo = float(costo_str)
            margen_input = float(self.entry_margen.get())
            margen_decimal = margen_input / 100
            
            precio_venta = costo / (1 - margen_decimal)
            ganancia = precio_venta - costo

            self.label_precio.configure(text=f"PRECIO VENTA: ${self.formato_moneda(precio_venta)}")
            self.label_ganancia.configure(text=f"GANANCIA: ${self.formato_moneda(ganancia)}")

            datos = {
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Costo": costo,
                "Precio Venta": round(precio_venta, 2),
                "Ganancia": round(ganancia, 2),
                "Margen %": margen_input
            }

            archivo = "registro_precios.xlsx"
            if not os.path.isfile(archivo):
                pd.DataFrame([datos]).to_excel(archivo, index=False)
            else:
                df_existente = pd.read_excel(archivo)
                # IMPORTANTE: El nuevo dato se concatena AL FINAL (debajo)
                df_nuevo = pd.concat([df_existente, pd.DataFrame([datos])], ignore_index=True)
                df_nuevo.to_excel(archivo, index=False)

            self.actualizar_tabla()
            self.entry_costo.delete(0, 'end')
            self.entry_costo.focus()
        except: pass
        
        self.tabla.configure(state="disabled")

    def limpiar_historial(self):
        archivo = "registro_precios.xlsx"
        if os.path.isfile(archivo):
            os.remove(archivo)
        self.actualizar_tabla()

    def procesar(self):
        try:
            costo_str = self.entry_costo.get().replace(".", "").replace(",", "")
            costo = float(costo_str)
            margen_input = float(self.entry_margen.get())
            margen_decimal = margen_input / 100
            
            precio_venta = costo / (1 - margen_decimal)
            ganancia = precio_venta - costo

            self.label_precio.configure(text=f"PRECIO VENTA: ${self.formato_moneda(precio_venta)}")
            self.label_ganancia.configure(text=f"GANANCIA: ${self.formato_moneda(ganancia)}")

            datos = {
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Costo": costo,
                "Precio Venta": round(precio_venta, 2),
                "Ganancia": round(ganancia, 2),
                "Margen %": margen_input
            }

            archivo = "registro_precios.xlsx"
            if not os.path.isfile(archivo):
                pd.DataFrame([datos]).to_excel(archivo, index=False)
            else:
                df = pd.concat([pd.read_excel(archivo), pd.DataFrame([datos])], ignore_index=True)
                df.to_excel(archivo, index=False)

            self.actualizar_tabla()
            self.entry_costo.delete(0, 'end')
            self.entry_costo.focus()
        except: pass

if __name__ == "__main__":
    app = CalculadoraALEXSolutions()
    app.mainloop()
