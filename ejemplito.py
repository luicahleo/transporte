# -*- coding: utf-8 -*-
"""
Created on Fri Nov  5 10:29:38 2021

@author: Gabriel
"""
from logging import info

from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *

name='TRANSB.xlsx'
excel_doc=openpyxl.load_workbook(name,data_only=True)
sheet=excel_doc['Hoja1']

arcos=Read_Excel_to_List(sheet, 'B2', 'B14')
nodos=Read_Excel_to_List(sheet, 'F2', '75')
prod_dem=Read_Excel_to_List(sheet, 'G2', 'G7')
info_arcos=Read_Excel_to_NesteDic(sheet, 'B1', 'E14')


A={}

# construto matriz lleno de 0
for n in nodos:
    A[n]={}
    for a in arcos:
        A[n][a] = 0.0

# pongo 1  y 0 donde tengo que poner
for n in nodos:
    for a in arcos:
        if info_arcos[a]['i'] == n:
            A[n][a] = 1
        elif info_arcos[a]['j'] == n:
            A[n][a] = -1
        else:
            A[n][a] = 0




def ejemplito():
    solver=pywraplp.Solver.CreateSolver('GLOP')
    # definimos variables
    x={}
    rnodos={}

    for a in arcos:
        # a es un caracater no un numero entero
        x[a] = solver.NumVar(0,solver.infinity(),'X%s'%(a))

    print('Número de variables=',solver.NumVariables())

    # tenemos que hacer balance de nodos
    for n in nodos:
        rnodos[n]=solver.Add(sum(A[n][a]*x[a] for a in arcos)==prod_dem[n-1],'RF%d'%(n))
        
    print('Número de restricciones=',solver.NumConstraints())

    # minimizar la suma de infor_arcos
    solver.Minimize(solver.Sum(info_arcos[a]['coste']*x[a] for a in arcos))
    
    status= solver.Solve()
    
    if status==pywraplp.Solver.OPTIMAL:
    #     for i in Fabricas:
    #         for j in Almacenes:
    #             print('X%d;%d = %d' %
    #                       (i,j,x[i][j].solution_value()))
    #     for i in Fabricas:
    #         for j in Almacenes:
    #             print('CR%d;%d = %d' %
    #                       (i,j,x[i][j].ReducedCost()))
    #     for i in Fabricas:
    #         print('u%d=%d' %
    #                   (i,rfab[i].dual_value()))
    #     for j in Almacenes:
    #         print('v%d=%d' %
    #                   (j,ralm[j].dual_value()))
        print('Función objetivo =', solver.Objective().Value())
    else:
        print('El problema es inadmisible')
    
    Solu={}
    for a in arcos:
        Solu[a]=0.0
    
    for a in arcos:
        Solu[a]=x[a].solution_value()
            
    Write_DicTable_to_Excel(excel_doc, name, sheet, Solu, 'I2', 'I14')
    
    
ejemplito()
    
              



      
            
            