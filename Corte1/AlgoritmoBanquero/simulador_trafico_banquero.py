import random
import threading
import time

# -----------------------------------------------------------------------------------
# 1. CONFIGURACION Y DATOS DEL PROBLEMA
# -----------------------------------------------------------------------------------

# Escala de tiempo: baja este numero para que la simulacion corra mas rapido.
# Con TIME_SCALE = 1.0 se respetan los 30 segundos reales por vehiculo.
TIME_SCALE = 0.05

TIEMPO_CRUCE = 30.0 * TIME_SCALE      # segundos que tarda 1 auto en cruzar
TIEMPO_AMARILLO = 5.0 * TIME_SCALE    # segundos que dura la luz amarilla

# Si esta en True, se imprimen tambien los pasos internos del Banquero.
MOSTRAR_BANQUERO = True

# Cuantos autos tiene cada sentido (dato del enunciado).
AUTOS = {
    "A_to_B": 50,
    "B_to_A": 35,
    "C_to_D": 20,
    "D_to_C": 100,
}

# A que eje pertenece cada sentido.
EJE_DE = {
    "A_to_B": "H",
    "B_to_A": "H",
    "C_to_D": "V",
    "D_to_C": "V",
}

# Que sentidos hay en cada eje (lo contrario del diccionario anterior).
SENTIDOS_DEL_EJE = {
    "H": ["A_to_B", "B_to_A"],
    "V": ["C_to_D", "D_to_C"],
}

# Recursos totales del Banquero: 2 instancias por eje (una para cada sentido
# de ese eje, asi los dos pueden circular juntos).
RECURSOS_TOTALES = {"H": 2, "V": 2}

# Hora de inicio de la simulacion, para calcular "Tiempo: Ys" en cada linea.
tiempo_inicio = time.time()

# Lock solo para que los print() de distintos hilos no se mezclen entre si.
lock_impresion = threading.Lock()


def imprimir_semaforo(sentido, estado, restantes):
    """Imprime una linea con el formato pedido:
    [Tiempo: Ys] [Semaforo Sentido X] Estado: [VERDE/AMARILLO/ROJO] | Carros restantes: Z
    """
    segundos = round(time.time() - tiempo_inicio, 2)
    lock_impresion.acquire()
    try:
        print("[Tiempo: " + str(segundos) + "s] [Semaforo " + sentido + "] "
              "Estado: " + estado + " | Carros restantes: " + str(restantes))
    finally:
        lock_impresion.release()


def imprimir_banquero(mensaje):
    """Imprime una linea de traza del Algoritmo del Banquero."""
    if not MOSTRAR_BANQUERO:
        return
    segundos = round(time.time() - tiempo_inicio, 2)
    lock_impresion.acquire()
    try:
        print("[Tiempo: " + str(segundos) + "s]" + mensaje)
    finally:
        lock_impresion.release()


# -----------------------------------------------------------------------------------
# 2. ESTRUCTURAS DEL ALGORITMO DEL BANQUERO (variables globales, sin clases)
# -----------------------------------------------------------------------------------

# Available[recurso] = instancias libres en este momento.
disponible = dict(RECURSOS_TOTALES)

# Max[sentido][recurso] = lo maximo que ese sentido puede llegar a pedir.
# Cada sentido solo necesita 1 unidad de SU propio eje y 0 del otro.
maximo = {}
for sentido in EJE_DE:
    maximo[sentido] = {}
    for recurso in ["H", "V"]:
        if EJE_DE[sentido] == recurso:
            maximo[sentido][recurso] = 1
        else:
            maximo[sentido][recurso] = 0

# Allocation[sentido][recurso] = lo que tiene asignado ahora mismo (empieza en 0).
asignado = {}
for sentido in EJE_DE:
    asignado[sentido] = {"H": 0, "V": 0}

# Need[sentido][recurso] = Max - Allocation (al inicio, Need = Max).
necesidad = {}
for sentido in EJE_DE:
    necesidad[sentido] = dict(maximo[sentido])

# Condition = un Lock + capacidad de wait()/notify_all(). Toda la logica del
# Banquero (leer y modificar disponible/asignado/necesidad) ocurre dentro de
# "with condicion:" para que sea atomica y no haya condiciones de carrera.
condicion = threading.Condition()


def eje_ocupado(eje):
    """Devuelve True si algun sentido de ese eje tiene recursos asignados
    en este momento (es decir, hay trafico circulando en ese eje)."""
    for sentido in SENTIDOS_DEL_EJE[eje]:
        if asignado[sentido][eje] > 0:
            return True
    return False


def eje_contrario(eje):
    if eje == "H":
        return "V"
    else:
        return "H"


def es_estado_seguro():
    """Algoritmo clasico de seguridad del Banquero, escrito con listas y
    diccionarios simples (sin numpy ni estructuras avanzadas).

    Simula, con un vector de trabajo (work) y uno de finalizacion (finish),
    si TODOS los procesos podrian terminar en algun orden dado el estado
    actual. Si es asi, el estado es seguro (no puede derivar en deadlock).
    """
    trabajo = dict(disponible)
    terminado = {}
    for sentido in EJE_DE:
        terminado[sentido] = False

    hubo_avance = True
    while hubo_avance:
        hubo_avance = False
        for sentido in EJE_DE:
            if terminado[sentido]:
                continue
            # Puede terminar si su necesidad cabe en lo disponible ahora mismo.
            puede_terminar = True
            for recurso in ["H", "V"]:
                if necesidad[sentido][recurso] > trabajo[recurso]:
                    puede_terminar = False
            if puede_terminar:
                # Simulamos que termina: devuelve lo que tenia asignado.
                for recurso in ["H", "V"]:
                    trabajo[recurso] = trabajo[recurso] + asignado[sentido][recurso]
                terminado[sentido] = True
                hubo_avance = True

    # El estado es seguro solo si TODOS los procesos pudieron terminar.
    for sentido in EJE_DE:
        if not terminado[sentido]:
            return False
    return True


def pedir_recurso(sentido):
    """Pide, para 'sentido', 1 unidad del recurso de su eje. Si no se puede
    conceder todavia, el hilo se BLOQUEA (condicion.wait()) sin consumir CPU,
    hasta que otro hilo libere recursos y lo despierte (notify_all)."""
    eje = EJE_DE[sentido]
    otro_eje = eje_contrario(eje)

    condicion.acquire()
    try:
        while True:
            imprimir_banquero(sentido + " solicita el recurso '" + eje + "'.")

            # Regla de exclusion mutua entre autopistas: si el otro eje esta
            # ocupado, este sentido debe esperar (semaforo en ROJO).
            if eje_ocupado(otro_eje):
                imprimir_banquero("Eje '" + otro_eje + "' ocupado -> " +
                                   sentido + " espera (ROJO).")
                condicion.wait()
                continue

            if disponible[eje] < 1:
                imprimir_banquero("No hay unidades libres de '" + eje +
                                   "' para " + sentido + ". Esperando...")
                condicion.wait()
                continue

            # --- Asignacion TENTATIVA (paso obligatorio del Banquero) ---
            disponible[eje] = disponible[eje] - 1
            asignado[sentido][eje] = asignado[sentido][eje] + 1
            necesidad[sentido][eje] = necesidad[sentido][eje] - 1

            if es_estado_seguro():
                imprimir_banquero("Prueba de estado seguro: OK. Recurso '" +
                                   eje + "' otorgado a " + sentido + ".")
                condicion.notify_all()
                return  # Se concede el recurso: el hilo puede poner VERDE.

            # Si el estado quedara inseguro, se revierte la asignacion.
            disponible[eje] = disponible[eje] + 1
            asignado[sentido][eje] = asignado[sentido][eje] - 1
            necesidad[sentido][eje] = necesidad[sentido][eje] + 1
            imprimir_banquero("Prueba de estado seguro: FALLO para " +
                               sentido + ". Se revierte la asignacion.")
            condicion.wait()
    finally:
        condicion.release()


def liberar_recurso(sentido):
    """Libera el recurso de eje que 'sentido' tenia asignado (el auto ya
    cruzo) y avisa a todos los hilos que puedan estar esperando."""
    eje = EJE_DE[sentido]
    condicion.acquire()
    try:
        disponible[eje] = disponible[eje] + 1
        asignado[sentido][eje] = asignado[sentido][eje] - 1
        necesidad[sentido][eje] = necesidad[sentido][eje] + 1
        imprimir_banquero(sentido + " libera el recurso '" + eje +
                           "'. Disponible[" + eje + "] = " + str(disponible[eje]))
        condicion.notify_all()
    finally:
        condicion.release()


def imprimir_matrices_iniciales():
    """Muestra en consola el estado inicial de Allocation, Max y Need."""
    print("")
    print("=" * 85)
    print("ESTADO INICIAL DE LAS ESTRUCTURAS DEL ALGORITMO DEL BANQUERO")
    print("=" * 85)
    print("Recursos totales:     " + str(RECURSOS_TOTALES))
    print("Recursos disponibles: " + str(disponible))
    print("")
    print("Proceso     Allocation           Max                  Need")
    print("-" * 60)
    for sentido in EJE_DE:
        print(sentido.ljust(12) + str(asignado[sentido]).ljust(21) +
              str(maximo[sentido]).ljust(21) + str(necesidad[sentido]))
    print("=" * 85)
    print("")


# -----------------------------------------------------------------------------------
# 3. FUNCION QUE EJECUTA CADA HILO (uno por sentido de circulacion)
# -----------------------------------------------------------------------------------

def simular_sentido(sentido):
    """Logica de un hilo: mientras le queden autos, pide el recurso al
    Banquero, pone VERDE, cruza un auto, pone AMARILLO, libera el recurso
    y vuelve a ROJO. Al terminar todos sus autos, hace una pausa final."""
    restantes = AUTOS[sentido]
    imprimir_semaforo(sentido, "ROJO", restantes)

    while restantes > 0:
        # 1) Pedir el recurso (puede bloquear el hilo hasta que sea seguro).
        pedir_recurso(sentido)

        # 2) VERDE: cruza un auto.
        imprimir_semaforo(sentido, "VERDE", restantes)
        time.sleep(TIEMPO_CRUCE)
        restantes = restantes - 1

        # 3) AMARILLO: aviso antes de ceder el paso.
        imprimir_semaforo(sentido, "AMARILLO", restantes)
        time.sleep(TIEMPO_AMARILLO)

        # 4) Se libera el recurso y vuelve a ROJO.
        liberar_recurso(sentido)
        imprimir_semaforo(sentido, "ROJO", restantes)

        # Pequena pausa para no acaparar el recurso apenas liberado.
        time.sleep(random.uniform(0.01, 0.05))

    imprimir_semaforo(sentido, "FINALIZADO", 0)
    time.sleep(0.5)  # suspension final del hilo, tal como pide el enunciado


# -----------------------------------------------------------------------------------
# 4. PROGRAMA PRINCIPAL
# -----------------------------------------------------------------------------------

def main():
    print("Iniciando simulacion de control de trafico en la interseccion...")
    print("TIME_SCALE = " + str(TIME_SCALE) +
          " | TIEMPO_CRUCE = " + str(round(TIEMPO_CRUCE, 2)) + "s" +
          " | TIEMPO_AMARILLO = " + str(round(TIEMPO_AMARILLO, 2)) + "s")

    imprimir_matrices_iniciales()

    hilos = []
    for sentido in AUTOS:
        hilo = threading.Thread(target=simular_sentido, args=(sentido,))
        hilos.append(hilo)

    for hilo in hilos:
        hilo.start()

    for hilo in hilos:
        hilo.join()

    total = round(time.time() - tiempo_inicio, 2)
    print("")
    print("=" * 85)
    print("SIMULACION FINALIZADA en " + str(total) + "s (todos los hilos terminaron)")
    for sentido in AUTOS:
        print("  - " + sentido + ": " + str(AUTOS[sentido]) + " vehiculos procesados.")
    print("=" * 85)


if __name__ == "__main__":
    main()