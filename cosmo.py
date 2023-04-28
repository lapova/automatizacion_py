import pandas as pd

#Lectura del archivo
preinscritos = pd.read_excel('preinscritos.xlsx')

#Se renombra la columna 'Sede' para estandarizacion de datos
preinscritos.rename(columns={'Sede': 'Centro de Experiencia'}, inplace=True)

#Se filtran todos los estudiantes cuyo estado es 'No continua el proceso'
preinscritos.drop(preinscritos[preinscritos['Estado de preinscripcion'] == 'No continúa el proceso'].index, inplace=True)


#Se cambia el nombre de todas los centro de experiencia para estandarizacion
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('bello'), 'Centro de Experiencia'] = 'Bello'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('belén'), 'Centro de Experiencia'] = 'Belén La Palma'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('envigado'), 'Centro de Experiencia'] = 'Envigado'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('rionegro'), 'Centro de Experiencia'] = 'Rionegro'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('perpetuo'), 'Centro de Experiencia'] = 'Perpetuo Socorro'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('centro'), 'Centro de Experiencia'] = 'Centro'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('barrio colombia'), 'Centro de Experiencia'] = 'Barrio Colombia'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('robledo'), 'Centro de Experiencia'] = 'Robledo'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('san vicente'), 'Centro de Experiencia'] = 'San Vicente'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('cristo rey'), 'Centro de Experiencia'] = 'Cristo Rey'
preinscritos.loc[preinscritos['Centro de Experiencia'].str.lower().str.contains('pérez'), 'Centro de Experiencia'] = 'Bello'


#Se cambian los nombres de los grados para el cruce correcto de informacion
preinscritos.loc[preinscritos['Grado'].str.lower().str.contains('cunas'), 'Grado'] = 'Sala cunas'
preinscritos.loc[preinscritos['Grado'].str.lower().str.contains('inspiratec'), 'Centro de Experiencia'] = 'Inspiratec'
preinscritos['Grado'] = preinscritos['Grado'].apply(lambda x: str(x).replace('InspiraTec', '').strip())
preinscritos.drop(preinscritos[preinscritos['Grado'].str.lower().str.contains('alcaldía')].index, inplace=True)
preinscritos.drop(preinscritos[preinscritos['Grado'].str.lower().str.contains('alcal')].index, inplace=True)
preinscritos.drop(preinscritos[preinscritos['Grado'].str.lower().str.contains('-')].index, inplace=True)

preinscritosCopy = preinscritos.copy()

#Se agrupan los estudiantes por centro de experiencia, grado y jornada
preinscritos = preinscritos.groupby(['Centro de Experiencia', 'Grado', 'Jornada', 'Estado de preinscripcion', 'Categoria de afiliación'])['Número de identificación'].count().reset_index(name='Conteo')
preinscritosCopy = preinscritosCopy.groupby(['Centro de Experiencia', 'Grado', 'Jornada', 'Estado de preinscripcion'])['Número de identificación'].count().reset_index(name='Conteo')

#Se crea la tabla dandole estructura con jornada y estado de preinscripcion
preinscritos = preinscritos.pivot_table(values='Conteo', index=['Centro de Experiencia', 'Grado'], columns=['Jornada', 'Categoria de afiliación', 'Estado de preinscripcion'])
preinscritosCopy = preinscritosCopy.pivot_table(values='Conteo', index=['Centro de Experiencia', 'Grado'], columns=['Jornada', 'Estado de preinscripcion'])

preinscritosCopy['AM'] = preinscritosCopy['AM'].fillna(0)
preinscritosCopy['PM'] = preinscritosCopy['PM'].fillna(0)

#Se eliminan los encabezados
preinscritos = preinscritos.rename_axis([None, None, None], axis=1)

#Se rellenan todos los vacios para que sean 0 y se puedan operar con ellos
preinscritos[('AM', 'A')] = preinscritos[('AM', 'A')].fillna(0)
preinscritos[('AM', 'B')] = preinscritos[('AM', 'B')].fillna(0)
preinscritos[('AM', 'C')] = preinscritos[('AM', 'C')].fillna(0)
preinscritos[('AM', 'D')] = preinscritos[('AM', 'D')].fillna(0)
preinscritos[('PM', 'A')] = preinscritos[('PM', 'A')].fillna(0)
preinscritos[('PM', 'B')] = preinscritos[('PM', 'B')].fillna(0)
preinscritos[('PM', 'C')] = preinscritos[('PM', 'C')].fillna(0)
preinscritos[('PM', 'D')] = preinscritos[('PM', 'D')].fillna(0)


with pd.ExcelWriter('preinscritosDinamica.xlsx') as writer:
    preinscritosCopy.to_excel(writer)

preinscripcionesX = pd.read_excel('preinscripciones.xlsx', sheet_name='Real vs Q10', header=3)

preinscripcionesX['Sede'] = preinscripcionesX['Sede'].fillna(method='ffill')

preinscripcionesX = preinscripcionesX.set_index(['Sede', 'Grado'])

preinscripcionesX.drop('Unnamed: 0', axis=1, inplace=True)

preinscripcionesX.rename(columns={'Unnamed: 2': None, 'Unnamed: 4': 'AM', 'Unnamed: 5': 'AM', 'Unnamed: 6': 'AM', 'Unnamed: 7': 'AM', 'Unnamed: 8': 'AM', 'Unnamed: 9': 'AM', 'Unnamed: 10': 'AM', 'Unnamed: 11': 'AM', 'Unnamed: 12': 'AM', 'Unnamed: 13': 'AM', 'Unnamed: 15': 'PM', 'Unnamed: 16': 'PM', 'Unnamed: 17': 'PM', 'Unnamed: 18': 'PM', 'Unnamed: 19': 'PM', 'Unnamed: 20': 'PM', 'Unnamed: 21': 'PM', 'Unnamed: 22': 'PM', 'Unnamed: 23': 'PM', 'Unnamed: 24': 'PM'}, inplace=True)

preinscripcionesX.columns = pd.MultiIndex.from_tuples([('AM', 'Capacidad Q10'), ('AM', 'Reserva Q10'),
    ('AM', 'Capacidad Real'), ('AM', 'Continuidad'), ('AM', 'Preinscritos - APROBADOS'),
    ('AM', 'Total cupos asignados'), ('AM', 'Admitidos Nuevos'), ('AM', 'Total Admitidos'),
    ('AM', 'Lista Espera'), ('AM', 'Disponibilidad Real'), ('AM', 'Disponibiilidad Q10'), ('PM', 'Capacidad Q10'),
    ('PM', 'Reserva Q10'), ('PM', 'Capacidad Real'), ('PM', 'Continuidad'), ('PM', 'Preinscritos - APROBADOS'),
    ('PM', 'Total cupos asignados'), ('PM', 'Admitidos Nuevos'), ('PM', 'Total Admitidos'),
    ('PM', 'Lista Espera'), ('PM', 'Disponibilidad Real'), ('PM', 'Disponibiilidad Q10')
])



for sede in preinscritos.index:
    aprobadosAM = preinscritosCopy.loc[sede, 'AM']
    aprobadosPM = preinscritosCopy.loc[sede, 'PM']
    if 'Sabaneta' not in str(sede):
        if 'nan' not in str(sede):
            preinscripcionesX.loc[sede, ('AM', 'Preinscritos - APROBADOS')] = aprobadosAM['Admitido'] + aprobadosAM['Aprobado']
            preinscripcionesX.loc[sede, ('AM', 'Lista Espera')] = aprobadosAM['En lista de espera']
            preinscripcionesX.loc[sede, ('PM', 'Preinscritos - APROBADOS')] = aprobadosPM['Admitido'] + aprobadosPM['Aprobado']
            preinscripcionesX.loc[sede, ('PM', 'Lista Espera')] = aprobadosPM['En lista de espera']


for index in preinscripcionesX.index:
    datosAM = preinscripcionesX.loc[sede, 'AM']
    datosPM = preinscripcionesX.loc[sede, 'PM']
    preinscripcionesX.loc[sede, ('AM', 'Total cupos asignados')] = datosAM['Continuidad'] + datosAM['Preinscritos - APROBADOS']
    preinscripcionesX.loc[sede, ('PM', 'Total cupos asignados')] = datosPM['Continuidad'] + datosPM['Preinscritos - APROBADOS']
    preinscripcionesX.loc[sede, ('AM', 'Total Admitidos')] = datosAM['Continuidad'] + datosAM['Admitidos Nuevos']
    preinscripcionesX.loc[sede, ('PM', 'Total Admitidos')] = datosPM['Continuidad'] + datosPM['Admitidos Nuevos']
    preinscripcionesX.loc[sede, ('AM', 'Disponibilidad Real')] = datosAM['Capacidad Real'] - datosAM['Total cupos asignados']
    preinscripcionesX.loc[sede, ('PM', 'Disponibilidad Real')] = datosPM['Capacidad Real'] - datosPM['Total cupos asignados']

with pd.ExcelWriter('informe.xlsx') as writer:
     preinscripcionesX.to_excel(writer)

index = preinscripcionesX.index

columns= pd.MultiIndex.from_tuples([
     # ----------------------------------- AM cat A -----------------------------------
    ('AM', 'A','Preinscritos - APROBADOS'), ('AM', 'A','Lista Espera'),

    # ----------------------------------- AM cat B -----------------------------------
    ('AM', 'B','Preinscritos - APROBADOS'), ('AM', 'B','Lista Espera'),

    # ----------------------------------- AM cat C -----------------------------------
    ('AM', 'C','Preinscritos - APROBADOS'), ('AM', 'C','Lista Espera'), 

    # ----------------------------------- AM cat D -----------------------------------
    ('AM', 'D','Preinscritos - APROBADOS'), ('AM', 'D','Lista Espera'), 

    # ----------------------------------- PM cat A -----------------------------------
    ('PM', 'A','Preinscritos - APROBADOS'), ('PM', 'A','Lista Espera'),

    # ----------------------------------- PM cat B -----------------------------------
    ('PM', 'B','Preinscritos - APROBADOS'), ('PM', 'B','Lista Espera'), 

    # ----------------------------------- PM cat C -----------------------------------
    ('PM', 'C','Preinscritos - APROBADOS'), ('PM', 'C','Lista Espera'),

    # ----------------------------------- PM cat D -----------------------------------
    ('PM', 'D','Preinscritos - APROBADOS'), ('PM', 'D','Lista Espera')
])

categorias = pd.DataFrame(
    columns=columns,
    index=index
)

for sede in preinscritos.index:
    # Categoria A
    aprobadosAMa = preinscritos.loc[sede, ('AM', 'A')]
    aprobadosPMa = preinscritos.loc[sede, ('PM', 'A')]
    # Categoria B
    aprobadosAMb = preinscritos.loc[sede, ('AM', 'B')]
    aprobadosPMb = preinscritos.loc[sede, ('PM', 'B')]
    # Categoria C
    aprobadosAMc = preinscritos.loc[sede, ('AM', 'C')]
    aprobadosPMc = preinscritos.loc[sede, ('PM', 'C')]
    # Categoria D
    aprobadosAMd = preinscritos.loc[sede, ('AM', 'D')]
    aprobadosPMd = preinscritos.loc[sede, ('PM', 'D')]
    if 'Sabaneta' not in str(sede):
        if 'nan' not in str(sede):
            # Categoria A
            categorias.loc[sede, ('AM', 'A', 'Preinscritos - APROBADOS')] = aprobadosAMa['Admitido'] + aprobadosAMa['Aprobado']
            categorias.loc[sede, ('AM', 'A', 'Lista Espera')] = aprobadosAMa['En lista de espera']
            categorias.loc[sede, ('PM', 'A', 'Preinscritos - APROBADOS')] = aprobadosPMa['Admitido'] + aprobadosPMa['Aprobado']
            categorias.loc[sede, ('PM', 'A', 'Lista Espera')] = aprobadosPMa['En lista de espera']
            # Categoria B
            categorias.loc[sede, ('AM', 'B', 'Preinscritos - APROBADOS')] = aprobadosAMb['Admitido'] + aprobadosAMb['Aprobado']
            categorias.loc[sede, ('AM', 'B', 'Lista Espera')] = aprobadosAMb['En lista de espera']
            categorias.loc[sede, ('PM', 'B', 'Preinscritos - APROBADOS')] = aprobadosPMb['Admitido'] + aprobadosPMb['Aprobado']
            categorias.loc[sede, ('PM', 'B', 'Lista Espera')] = aprobadosPMb['En lista de espera']
            # Categoria C
            categorias.loc[sede, ('AM', 'C', 'Preinscritos - APROBADOS')] = aprobadosAMc['Admitido'] + aprobadosAMc['Aprobado']
            categorias.loc[sede, ('AM', 'C', 'Lista Espera')] = aprobadosAMc['En lista de espera']
            categorias.loc[sede, ('PM', 'C', 'Preinscritos - APROBADOS')] = aprobadosPMc['Admitido'] + aprobadosPMc['Aprobado']
            categorias.loc[sede, ('PM', 'C', 'Lista Espera')] = aprobadosPMc['En lista de espera']
            # Categoria D
            categorias.loc[sede, ('AM', 'D', 'Preinscritos - APROBADOS')] = aprobadosAMd['Admitido'] + aprobadosAMd['Aprobado']
            categorias.loc[sede, ('AM', 'D', 'Lista Espera')] = aprobadosAMd['En lista de espera']
            categorias.loc[sede, ('PM', 'D', 'Preinscritos - APROBADOS')] = aprobadosPMd['Admitido'] + aprobadosPMd['Aprobado']
            categorias.loc[sede, ('PM', 'D', 'Lista Espera')] = aprobadosPMd['En lista de espera']

with pd.ExcelWriter('informeCategorias.xlsx') as writer:
     categorias.to_excel(writer)

print('proceso completado')
