# EUROS_EN_LETRAS para Excel (VBA)

Convierte cantidades numéricas a **texto en euros** para facturas en Excel.
- `100` → **CIEN EUROS**
- `1234,50` → **MIL DOSCIENTOS TREINTA Y CUATRO EUROS CON CINCUENTA CÉNTIMOS**
- `100,00` → **CIEN EUROS** (si no hay céntimos, no los muestra)

> Compatible con **Excel de escritorio** (Windows/Mac). **No** funciona en Excel Online ni en la “app nueva” sin VBA.

---

## 📦 Archivos
- `src/EUROS_EN_LETRAS.bas` → módulo VBA listo para importar.
- `README.md` → este documento.

## ✅ Requisitos
- Excel de escritorio (Office/Microsoft 365 clásico).
- Guardar el libro como **.xlsm** (habilitado para macros).
- Macros habilitadas al abrir el archivo.

## 🛠️ Instalación (paso a paso)
1. Abre tu libro de Excel de facturas.
2. Pulsa **ALT+F11** (Mac: **⌥ Option + F11**) para abrir el Editor de VBA.
3. Menú **Archivo > Importar archivo…** → selecciona `src/EUROS_EN_LETRAS.bas`.
4. Menú **Depurar > Compilar VBAProject** (opcional, para comprobar errores).
5. Guarda el libro como **.xlsm** y, al volver a Excel, **habilita macros** si aparece la barra amarilla.

## 🧪 Uso
En una celda, escribe:
```
=EUROS_EN_LETRAS(A1)
```
Donde `A1` contiene el importe (número con decimales).  
- Si hay céntimos: *“… EUROS CON … CÉNTIMOS”*  
- Si no hay céntimos: *“… EUROS”*

> La función devuelve el texto **en mayúsculas**. Si lo prefieres en minúsculas:
> ```
> =MINUSC(EUROS_EN_LETRAS(A1))
> ```

## 🗓️ Frase de fecha automática estilo factura
**Fecha de hoy (automática):**
```
="MÁLAGA, a " & TEXTO(HOY();"[$-es-ES]d ""de"" mmmm ""del"" yyyy")
```
**Usando la fecha de la factura en B2:**
```
="MÁLAGA, a " & TEXTO(B2;"[$-es-ES]d ""de"" mmmm ""del"" yyyy")
```
Tips:
- Mes en mayúsculas: usa `MAYUSC(TEXTO(...))`.
- Si prefieres “de 2025” en lugar de “del 2025”, cambia `" del "` por `" de "`.

## ℹ️ Detalles técnicos
- Tipo **Currency** y redondeo a 2 decimales estilo Excel (half-up).
- Soporta importes habituales de facturación.
- Ajusta correctamente **“un euro / veintiún euros”** (apócope).
- Idioma fijo **es-ES** para los meses con `[$-es-ES]`.

## 🧯 Problemas comunes
- **#¿NOMBRE?** → o bien las macros están deshabilitadas, o pegaste el código en la *Hoja* en vez de en un **Módulo**.
- **“Error de compilación: Sub o Function no definida”** → el módulo no está completo; vuelve a importar `EUROS_EN_LETRAS.bas`.
- **ALT+F11 no abre** → probablemente estás en **Excel Online** o la **app nueva** sin soporte VBA. Usa Excel de escritorio.

## 🤝 Contribuir
- Acepta PRs y *issues* con mejoras, bugs o variantes (p. ej., céntimos en formato `00/100`).

## 📄 Licencia
MIT. Puedes usar, copiar y modificar libremente manteniendo el aviso de licencia.

---

Autor: Javier Ramírez Bandera
