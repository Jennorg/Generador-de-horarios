# -*- coding: utf-8 -*-

class Materia:

    def __init__(self,nombre,carrera,semestre,dia,hora_inicio,hora_final):
        self.nombre=nombre
        self.semestre=semestre
        self.carrera=carrera
        self.dia=dia
        self.hora_inicio=hora_inicio
        self.hora_final=hora_final
    
    def imprimir(self):
        print('Nombre: {}, Semestre: {}, carrera:{}, Dia: {} , Hora de inicio:{}, Hora de Finalizacion:{}'.format(self.nombre, self.semestre,self.carrera, self.dia, self.hora_inicio,self.hora_final))
    

