#!/usr/bin/env python
# coding: utf-8

# # Introducción a python para Machine Learning
# En este cuaderno exploraremos rápidamente el lenguaje de programación Python. No se entrará en profundidad, pero se mostrará todo lo que hay por aprender. 
# 
# Todo lo que está acá se puede ampliar en:
# 
# 
# *   https://www.tutorialspoint.com/python3/index.htm
# *   https://www.tutorialspoint.com/numpy/
# *   https://www.tutorialspoint.com/python_pandas/index.htm
# *   https://www.google.com/
# 
# 
# 
# **Nota:** recordar que usaremos Python 3.6

# **Instalación:**
# * Lo más fácil: instalar Anaconda (https://anaconda.org/anaconda/python)
# * Lo recomendado: instalar Python (https://www.python.org/) y usar [Virtualenv](https://virtualenv.pypa.io/en/stable/)
# 
# Para más información, consulte [esta guía](https://drive.google.com/open?id=1aIaArcdHWusDfGmO4BvRzBDoWH8RSPr-) (advertencia: tiene varios errores de ortografía) o cualquier otra guía en internet.

# ## Lo estándar

# **Imprimir mensajes**

# In[3]:


print("Este es un mensaje para saludar al mundo: hola mundo!")


# In[5]:


msg = "También puedes guardar mensajes en una variable para luego imprimirlos"
print(msg)


# ** Operaciones matemáticas**

# In[ ]:


a = 3
b = 8


# multiplicación

# In[9]:


a*b


# potenciación

# In[10]:


a**b


# división

# In[11]:


a/b


# módulo: si divido b entre a, cuánto me sobra?

# In[12]:


b%a


# cociente: cuántas partes enteras de a caben en b

# In[13]:


b//a


# Hay muchas más operaciones matemáticas. Puedes revisar el módulo [math](https://docs.python.org/3/library/math.html)

# In[ ]:


import math


# In[ ]:


math.sqrt(b)


# **Estructuras básicas**

# Listas

# In[ ]:


[2, 6, 7, 8]


# In[ ]:


['a', 'b', 'c']


# Las listas pueden guardar cualquier cosa

# In[ ]:


[3, None, int, str, "casa", math.pi]


# incluso otras listas

# In[ ]:


[3, [4,5,6], 7, [8,9]]


# A las listas se les pueden agregar elementos en cualquier momento

# In[ ]:


lista = ['a', 'b', 'c', 'd', 'd', 'd']


# In[ ]:


lista.append('otro caracter')
lista


# Se puede acceder a los elementos de una lista basado en el índice

# In[ ]:


lista[0]


# In[ ]:


lista[1]


# In[ ]:


lista[-1]


# In[ ]:


lista[:2]


# In[ ]:


lista[2:]


# In[ ]:


lista[2:4]


# In[ ]:


lista[2] = 25


# In[17]:


lista


# también podemos preguntar por el índice de un elemento particular de la lista

# In[ ]:


lista.index('d')


# adicionalmente, es posible pasar de un string con alguna separación entre elementos a una lista (piensa en un csv):

# In[15]:


lista = "nombre, apellido, ciudad, animal, fruta, color, cosa".split(',')
lista


# incluso podemos comparar las listas elemento por elemento

# In[ ]:


['nombre', ' apellido', ' ciudad', ' animal', ' fruta', ' color', ' cosa'] == lista


# In[ ]:


['nombre', ' apellido', ' ciudad', ' animal', ' fruta', ' color', ' cosota'] == lista


# **tuplas**
# 
# son casi como las listas, solo que no se pueden modificar

# In[ ]:


tupla = ('casa', 'castillo', 400)
tupla


# In[ ]:


tupla[0] = 20


# **diccionarios**
# 
# Sirven para guardar valores (que pueden ser cualquier cosa: strings, números, listas, otros diccionarios, etc) y buscarlos utilizando una llave (que solo pueden ser valores inmutables).
# 

# In[18]:


d = {'a':1, 'b':2, 'c':3}
d


# In[23]:


var = 'o'
d = {'a':1, 'b':2, 'c':3, var: 3}
d


# In[31]:


d = {'a':1, 'b':2, 'c':3, 'a': 3}
d


# In[24]:


d['b']


# Están diseñados para ser accedidos mediante la llave, no se puede usar un índice; incluso el orden en el que se guardan no se garantiza.

# In[25]:


d[0]


# In[29]:


d = {('a','b'): 3, ('c', 'd'): 7, 'otro': 9}
d


# In[27]:


d[('a', 'b')]


# In[ ]:


d = {['a','b']: 3, ('c', 'd'): 7, 'otro': 9}


# **Condicionales**

# In[34]:


print(a, b)


# In[32]:


if b > a:
    print("b({}) es mayor que a({})".format(b, a))


# In[ ]:


if a > b:
    print("a({}) es mayor que b({})".format(a, b))


# In[35]:


if a > b:
    print("a({}) es mayor que b({})".format(a, b))
else:
    print("a({}) NO es mayor que b({})".format(a, b))


# **Loops**
# 
# Muchos objetos o typos de datos (de la librería estándar) se pueden iterar.
# 
# Iterar sobre strings

# In[37]:


super_string = "abcdef"
for ch in super_string:
    print(ch)


# Iterar sobre listas

# In[40]:


lista = ['Mi', 'nombre', 'es', 3]
for el in lista:
    print(el)


# Algunos estarán familiarizados con la forma de iterar en otros lenguajes, acá les muestro cómo:

# In[ ]:


lista = ['Mi', 'nombre', 'es', 3]
for i in range(len(lista)):
    print(lista[i])


# pero la forma de python es definitivamente más chévere, sin contar con que en algunos casos esa forma no funcionaría, por ejemplo cuando se itera en diccionarios.

# In[ ]:


d = {'a':1, 'b':2, 'c':3}
for k in d:
    print("key: {}, value: {}".format(k, d[k]))


# In[43]:


d = {'a':1, 'b':2, 'c':3}
for k in d.keys():
    print("key: {}, value: {}".format(k, d[k]))


# In[42]:


d.items()


# In[ ]:


d = {'a':1, 'b':2, 'c':3}
for k, v in d.items():
    print("key: {}, value: {}".format(k, v))


# In[48]:


d = {'a':1, 'b':2, 'c':3}
for v in d.values():
    print("value: {}".format(v))
print("esto solo va una vez")


# In[44]:


for i in range(len(d)):
    print(i)


# In[45]:


d[0]


# In[ ]:


d = {'a':1, 'b':2, 'c':3}
for i in range(len(d)):
    print(d[i])


# **while**

# In[46]:


a = 0
while a <= 3:
    print(a)
    a += 1


# In[47]:


a = 0
while a <= 3:
    print(a)
    a = a + 1


# **funciones**
# 
# Sin valores por defecto

# In[54]:


def una_funcion(a, b):
    print("el valor de a es {}:\nel valor de b es {}".format(a,b))
    
una_funcion(3,4)
una_funcion(None,50)
una_funcion('casa', 1)


# Con valores por defecto

# In[55]:


def una_funcion(a, b=3):
    print("el valor de a es {}:\nel valor de b es {}".format(a,b))
    
una_funcion(3,4)
una_funcion(None,50)
una_funcion('casa')
una_funcion('amarillo')


# In[64]:


def fn(a,b,c):
    print(a,b,c)
    
fn(3, 4, c=9)


# In[2]:


def fn(a,b,c,d,e,f,g,h):
    print(a,h)
l = [1,2,3,4,5,67,7,4]
fn(*l)


# In[1]:


def fn(a, *args):
    print(a, args)
    
fn(3, 'casa', 45)


# In[4]:


def fn(a, **kwargs):
    flag = kwargs.get('flag', None)
    
    if flag:
        pass
    else:
        pass
    print(a, kwargs)
    
fn(2)


# In[6]:


def fn (edad, nombre='nn', apellido='nn'):
    print(edad, nombre, apellido)

dic = {'nombre': 'felipe'}

fn(34, **dic)


# In[83]:


def fn(a, b=4, **kwargs):
    print(a, b, kwargs)

b=99
fn(70, c=b, b=40)


# In[7]:


def fn(a, b, c=4, d=3, *tuplaSobrante, **kwargs):
    print(a,b,c,d,tuplaSobrante,kwargs)
    
fn(3, 5, 3,4,5,6,'casa', nueve=6, cuarenta=7)


# In[36]:


def fn(a, b, c ,d):
    def fn_interna (e):
        return a*e
    return fn_interna


# In[39]:


fn(2,3,5,2)(8)


# In[35]:


dd


# **Clases y objetos**

# In[11]:


class Mano:
    material = 'carne'
    def __init__(self, n_dedos=5, color='piel'):
        self.n_dedos = n_dedos
        self.color = color
        self.temp_dedos = n_dedos
        
    def iscomplete(self):
        if self.n_dedos == 5:
            return True
        else:
            return False
        
    def cerrar(self):
        self.n_dedos = 0
        
    def abrir(self):
        self.n_dedos = self.temp_dedos


# In[16]:


mano1 = Modelo(4, 'morada')
mano1.entrenar()
mano1.predecir()


# In[19]:


mano1.abrir()


# In[21]:


a = 9


# In[20]:


mano1.n_dedos


# In[91]:


mano = Mano()
mano.iscomplete()


# In[92]:


mano = Mano(n_dedos=3)
mano.iscomplete()


# ## Lo estándar 2.0
# 

# **Contar elementos retornados durante el loop:** sirve para saber cuántos elementos e visto en el for loop y o para conocer el índice del elemento al que se accede en cada iteración

# In[22]:


words = "nombre, apellido, ciudad, animal, fruta, color, cosa".split(',')

for i, wd in enumerate(words):
    print("word {} is in index {}".format(wd, i))


# In[ ]:





# **Compresión de listas**
# ver: https://docs.python.org/3/tutorial/datastructures.html#list-comprehensions
# 
# Supongamos que dada una lista de palabras queremos sacar una lista que nos de la longitud de cada palabra

# In[41]:


words = "nombre, apellido, ciudad, animal, fruta, color, cosa".split(',')


# opcion 1

# In[42]:


words_length = []
for wd in words:
    if len(wd) != 7:
        words_length.append(len(wd))
    
words_length


# opción 2 (usualmente es más rápida)

# In[44]:


words_length = [len(wd) for wd in words if len(wd) != 7]
words_length


# esta otra opción tiene más opciones

# In[100]:


words

# In[]

#words[0]
len(words[0])%2

# In[98]:

words_length = [len(wd) for wd in words if len(wd)%2==True]
# Es igual a esto: 
#words_length = [len(wd) for wd in words if len(wd)%2]

words_length


# In[54]:


flag = False
word = 'casa' if flag else 'edificio'


# In[55]:


word


# In[50]:


[wd.strip().upper() for wd in words if wd.startswith(' ')]


# In[ ]:


words_length = [len(wd) if len(wd)%2 else 0 for wd in words]
words_length


# La misma idea es aplicable a diccionarios

# In[57]:

d = {k:v+1 for v,k in enumerate('abcdefghij')}
d


# Lectura recomendada: https://realpython.com/introduction-to-python-generators/

# **Excepciones**

# In[58]:


words


# In[59]:


words[200]


# In[61]:


try:
    val = words[200]
except IndexError as e:
    print("la lista es muy pequeña")
    val = words[-1]

# In[]
val
words

# In[66]:


raise Exception('el modulo del metodo tal necesita que primero se llame a la bse datos')


# Lectura recomendada: https://realpython.com/python-exceptions/

# **sets:** permiten, entre otras cosas, averiguar si un conjunto de elementos está presente en otros conjunto.
# 
# Lectura recomendada: https://docs.python.org/3/library/stdtypes.html#set
# 

# In[67]:


ll = [1,1,2,3,4,4,4,4,5,6,7,87,8,8,8]
list(set(ll))


# In[69]:


{'a', 'b'}.issubset({ch for ch in 'abcdefjhijklmnñopqrstuvwxyz'})


# In[70]:


{ch for ch in 'abcdefjhijklmnñopqrstuvwxyz'}.difference({'a', 'b'})


# **Args y Kwargs:** https://pythontips.com/2013/08/04/args-and-kwargs-in-python-explained/

# ## Numpy

# In[73]:


import numpy as np


# In[75]:


x = np.array([1,2]) #vector
y = np.matrix([[1,2],[4,5]]) #matriz


# In[76]:


x


# In[77]:


x.shape


# In[79]:


y.shape


# en generarl se usa `np.array` para cualquier arreglo multidimensional

# In[116]:


y = np.array([[1,2],[4,5]]) #matriz
y


# In[117]:


y.shape


# dimensiones del arreglo

# In[ ]:


x.shape


# In[ ]:


y.shape


# se puede calcular la transpuesta de un arreglo muy fácilmente

# In[80]:


y


# In[81]:


y.T


# y la inversa

# In[82]:


np.linalg.inv(y)


# multiplicación de matrices

# In[119]:


x.shape


# In[120]:


y.shape


# In[84]:


A = np.array([[1,2,3], [4,5,6]])
B = np.array([[1,2], [3,4]])
print(A.shape, B.shape)


# In[85]:


B*A


# In[86]:


np.dot(B, A)


# In[87]:


np.dot(x, y)


# También se puede sumar por filas o por columnas de un arreglo

# In[93]:


A


# In[92]:


np.sum(A, axis=0)


# In[ ]:


np.sum(y, axis=1)


# In[ ]:


scalar = np.random.randn()
scalar


# In[97]:


get_ipython().run_line_magic('pinfo', 'np.random.randn')


# In[ ]:


val = vector[-1]
val


# In[ ]:


vector[0] = 333
vector


# In[107]:


matrix = np.random.randn(3,5)
matrix


# In[104]:


matrix[1, 0] = 1
matrix


# In[124]:


matrix[:,2] = 5
matrix


# también podemos seleccionar elementos del arreglo usando un arreglo de 1s y 0s (el 1 para seleccionar elementos que queremos y el 0 para los que no queremos)

# In[109]:


matrix[[True, True, False]]


# In[110]:


matrix[[True, True, False],[False, False, True, True, False]]


# In[112]:


matrix


# También podemos hacer comparaciones con números y es como comparar cada elemento del arreglo con el número

# In[115]:


matrix[matrix >= 1]


# y usar esto para seleccionar todos los valores

# In[128]:


matrix[matrix >= 1]


# **Broadcasting:** sabemos que las operaciones matriciales tienen unas reglas estrictas, relacionadas con la dimensión de los arreglos, para poder realizarse. Broadcasting es una forma inteligente de interpretar las dimensiones de los arreglos antes de operarlos

# In[117]:


matrix


# In[118]:


matrix = matrix + 10
matrix


# In[119]:


matrix = matrix * 0.5
matrix


# In[121]:


matrix = matrix + np.array([1, 2, 3])
matrix


# In[123]:


matrix.shape


# In[126]:


np.array([1, 2, 3,4,5]).reshape(1,-1).shape


# In[125]:


matrix = matrix + np.array([1, 2, 3,4,5]).reshape(1,-1)
matrix


# In[127]:


matrix = matrix * np.array([-1, 0, 1, -1, 0])
matrix


# In[128]:


matrix = matrix * np.array([-1, 0, 1, -1, 0]).reshape(-1, 1)
matrix


# In[139]:


matrix * np.array([-1, 0, 1, -1, 0]).reshape(1, -1)


# cómo reversar el orden de las columnas?

# In[132]:


matrix[:,0:4:-1]


# In[138]:


matrix = matrix[...,::-1]
matrix


# arreglo de tres dimensiones

# In[ ]:


tensor = np.random.randn(2,3,4)
tensor


# In[ ]:


tensor.shape


# Este tensor se puede interpretar como cuatro matrices de $2\times 3$ cada una.
# 
# As que la primera de esas cuatro matrices se solo contenga 1

# In[139]:


matrix = np.random.randint(0, 10, [3,4])


# In[ ]:


tensor[...,0] = np.ones((2,3))
tensor


# si el resultado anterior no se ve como esperabas, no importa; eso no está hecho para verse

# También podemos reversar las capas

# In[ ]:


tensor = tensor[...,::-1]
tensor


# **concatenar**

# In[140]:


matrix


# supongamos que a la matriz anterior le queremos agregar una fila de 1s

# In[144]:


np.concatenate((matrix, np.ones((3,1)), np.zeros((3,3))), axis=1)


# o una columna

# In[ ]:


np.concatenate((matrix, np.ones((3,1))), axis=1)


# Hay muchas más cosas, por favor revisa: https://www.tutorialspoint.com/numpy/index.htm

# ## Pandas

# In[146]:


import pandas as pd


# In[ ]:


import os
os.listdir('sample_data')


# miremos el archivo california_housing_train.csv. Lo siguiente lo guarda en un objteto tipo pd.DataFrame

# In[147]:


get_ipython().run_line_magic('pinfo', 'pd.read_csv')


# In[ ]:


df = pd.read_csv('sample_data/california_housing_train.csv')


# la siguiente linea nos deja observar algunas lineas del archivo .csv

# In[ ]:


df.head()


# In[149]:


tabla = [{'col1': 1, 'col2': 2, 'col3': 3},{'col1': 1, 'col2': 6, 'col3': 3},{'col1': 5, 'col2': 2, 'col3': 4},{'col1': 5, 'col2': 2, 'col3': 3},{'col1': 9, 'col2': 2, 'col3': 0}]


# In[181]:


df = pd.DataFrame(tabla)
df


# In[158]:


df.loc[:2,'col1']


# In[160]:


df.iloc[0, 1]


# podemos sacar un rápido resumen de los datos usando pandas

# In[161]:


df.describe()


# incluso histogramas de cada variable (nota: hay mejores formas de sacar histogramas en Python)

# In[163]:


get_ipython().run_line_magic('matplotlib', 'inline')


# In[164]:


df.hist()


# se generar un dataframe con solo algunas columnas de interés

# In[169]:


df[['col1','col3']].head(100)


# In[168]:


df.tail()


# In[ ]:


df[['median_income', 'median_house_value']].head()


# In[189]:


df.loc[[0,3,4], ['col1','col2']] = 999
df.loc[[1,2], ['col3']] = -999
df


# In[190]:


df[(df >= 999) & (df <= -999)] = 5


# In[187]:


df


# Finalmente, acá hay muchas más cosas por revisar: https://www.tutorialspoint.com/python_pandas/index.htm

# ## Visualización
# 
# Acá no cubriremos esa parte, pero les recomiendo revisen estas librerías (no hace falta conocerlas todas, se las comparto por cultura general):
# 
# 
# *   https://matplotlib.org/
# *   https://plot.ly/
# *   https://bokeh.pydata.org/en/latest/
# *   https://seaborn.pydata.org/
# 
# 

# In[ ]:




