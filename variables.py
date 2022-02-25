import getpass

## Nombre de las variables agrupadas por Proceso
usuario_actual= getpass.getuser()

## Aqui va la ruta de los drivers
ruta_usuario_drivers = "TU RUTA"
ruta_usuario_drivers = fr"C:\Users\{usuario_actual}\Dropbox (Do Better)\Do Better's shared workspace\Desarrollo\recursos Soft\driver chrome\chromedriver.exe"

## Inciales
pagina_pbi = "https://powerbi.microsoft.com/es-es/"
xpath_inicio_sesion = '//*[@id="power-bi-portal-link-desktop"]'
xpath_input_usuario = '//*[@id="i0116"]'

## Usuario
correo = "TU@CORREO"
xpath_boton_siguiente = '//*[@id="idSIButton9"]'

## PASSWORD
xpath_input_pass = '//*[@id="i0118"]'
contra = "TU PASSWORD"
xpath_boton_inciar = '//*[@id="lightbox"]/div[3]/div/div[2]/div/div[4]/div[2]/div/div/div/div'

## quieres mantener la sesion inciada
xpath_boton_no = '//*[@id="idBtn_Back"]'

## URL DONDE ESTA LA BASE
url_base = "TU URL"

## Actualizacion
xpath_div_actualiza = '//*[@id="artifactContentList"]/div[1]/div[2]/div[2]'
xpath_actualiza = '//*[@id="artifactContentList"]/div[1]/div[2]/div[2]/span/button[1]/mat-icon'
