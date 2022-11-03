import re

def adicionar_seg(ruta:str) -> None:
    texto = ''
    patron_form_load = r"Private Sub Form_Load[^@]*?End Sub"
    try:
        with open(ruta, 'r') as form:
            texto = form.read()
            print('Original', texto)
            texto = reset_seguridad(texto)
            print(texto)
            texto = adicion_seguridad(texto)
            print(texto)
    except Exception as ex:
        print('Ha ocurrido un error:', ex)
    finally:
        print(f'Adicion de seguridad en {ruta}')

def reset_seguridad(texto:str) -> str:
    patron_seguridad = r"Call SeguridadSet(Me)"
    return re.sub(patron_seguridad, '', texto)

def adicion_seguridad(text:str) -> str:
    patron_load = r"Private Sub Form_Load[^@]*?End Sub"
    result = re.findall(patron_load, text)
    if len(result) == 0:
        return text
    cadena_anterior = result[0]
    cadena_nueva = re.sub(r'End Sub', r'\tCall SeguridadSet(Me)\nEnd Sub', cadena_anterior)
    return re.sub(cadena_anterior, cadena_nueva, text)

if __name__ == '__main__':
    adicionar_seg('frmLogin.frm')