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


# In[1]:


d = {'a':10, 'b':2, 'c':3, 'a': 3}
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

# In[14]:


a=8
b=13


# In[15]:


print(a, b)


# In[20]:


if b > a:
  print("b({}) es mayor que a({})".format(b, a))


# In[ ]:


if a > b:
    print("a({}) es mayor que b({})".format(a, b))


# In[35]:


if a > b:
    print("a({}) es mayor que b({})".format(a, b))
elif a==b:
    print("elif")
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


# In[68]:


def fn(a, *args):
    print(a, args)
    
fn(3, 'casa', 45)

