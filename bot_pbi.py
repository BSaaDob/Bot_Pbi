from index_pbi import *
from time import sleep
from variables import *

## Crear instancia 
bot_pbi = Actualiza_Reporte()

## Ingreso pagina
bot_pbi.driver.get(pagina_pbi)
bot_pbi.driver.maximize_window()

## Boton Inicio Sesion
bot_pbi.esperar_y_moverse( xpath= xpath_inicio_sesion)
bot_pbi.click_f()

sleep(5)
print("Espero")

## Definicion de pestanias para navegar entre ellas
window_before = bot_pbi.driver.window_handles[0]
window_after = bot_pbi.driver.window_handles[1]

## Cambio de pestania
bot_pbi.driver.switch_to.window( window_after)

## USUARIO 
bot_pbi.espera_y_escribe( xpath= xpath_input_usuario , texto= correo  )

bot_pbi.esperar_y_moverse( xpath= xpath_boton_siguiente)
bot_pbi.click_f()

## PASSWORD

bot_pbi.espera_y_escribe(xpath= xpath_input_pass , texto= contra)
bot_pbi.esperar_y_moverse( xpath= xpath_boton_inciar)
bot_pbi.click_f()
bot_pbi.click_f(tiempo=1)

## quieres mantener la sesion inciada
bot_pbi.esperar_y_moverse( xpath= xpath_boton_no)
bot_pbi.click_f()

## Moverse a la URL donde esta la base
bot_pbi.cambio_url( url= url_base)

## Actualizacion
bot_pbi.esperar_y_moverse( xpath= xpath_div_actualiza)
print("Me posicione")

## Posicionarse en el boton de actualizar
bot_pbi.esperar_y_moverse( xpath= xpath_actualiza)

## Bucle infinito
a = True
contador = 0
while a:
    sleep(30)
    bot_pbi.click_f()
    print(f"Esta es la iteracion Numero: {contador}")
    contador = contador + 1