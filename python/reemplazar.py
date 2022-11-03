import subprocess, sys, re

def obtener_frm():
    lista_rutas = []
    #subprocess.Popen(['powershell.exe', 'Remove-Item -Recurse -Force forms.txt'], stdout=sys.stdout)
    powershell_cmd = "Get-ChildItem -Recurse -Filter *.frm | Select FullName > forms.txt"
    subprocess.Popen(['powershell.exe', powershell_cmd], stdout=sys.stdout)
    with open('forms.txt', 'r', encoding='utf-16-le') as arch:
        for row in arch:
            if row.startswith('D'):
                lista_rutas.append(re.sub(r'\n', '', row))
    return lista_rutas
    
def adicionar_seg(ruta:str) -> None:
    texto = ''
    try:
        with open(ruta, 'r') as form:
            texto = form.read()
            texto = reset_seguridad(texto)
            texto = adicion_seguridad(texto)
        with open(ruta, 'w') as form_of:
            form_of.truncate()
            form_of.write(texto)
    except Exception as ex:
        print('Ha ocurrido un error:', ex)
    finally:
        print(f'Adicion de seguridad en {ruta}')

def reset_seguridad(texto:str) -> str:
    patron_seguridad = r"\n\tCall SeguridadSet\(Me\)"
    return re.sub(patron_seguridad, '', texto)

def adicion_seguridad(text:str) -> str:
    patron_load = r"Private Sub Form_Load[^@]*?End Sub"
    result = re.findall(patron_load, text)
    if len(result) == 0:
        return text
    cadena_anterior = result[0]
    cadena_nueva = re.sub(r'End Sub', r'\tCall SeguridadSet(Me)\nEnd Sub', cadena_anterior)
    return re.sub(patron_load, cadena_nueva, text)

def main():
    lista_rutas = obtener_frm()
    for ruta in lista_rutas:
        adicionar_seg(ruta)

if __name__ == '__main__':
    main()