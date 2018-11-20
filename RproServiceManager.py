# !python3
# RproServiceManager
# Version 1.1
# By Zach Cutberth

# Gui program to restart V9/Prism services.

# Imports
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import subprocess
import psutil
import os
import win32serviceutil
import win32service

services = ['PrismMQService',
            'PrismCommonService',
            'PrismResiliencyService',
            'PrismBackofficeService',
            'PrismV9Service',
            'PrismLicSvr',
            'RabbitMQ',
            'Apache',
            'MySQL57',
            'RProLicSvr',
            'OracleODS12cr1TNSListener',
            'OracleServiceRproODS']

dynamicButtons = []
dynamicLabels = []
dynamicCheckBoxes = []

def makeButtons():
    stopRow = 2
    startRow = 2
    restartRow = 2
    
    for service in services:
        
        stopButton = ttk.Button(allFrame, text='Stop', command=lambda service=service: stopStartRestart(service, 'stop'))
        dynamicButtons.append(stopButton)
        stopButton.grid(row=stopRow, column=3)
        stopRow += 1     
        
        startButton = ttk.Button(allFrame, text='Start', command=lambda service=service: stopStartRestart(service, 'start'))
        dynamicButtons.append(startButton)
        startButton.grid(row=startRow, column=4)
        startRow += 1     
        
        restartButton = ttk.Button(allFrame, text='Restart', command=lambda service=service: stopStartRestart(service, 'restart'))
        dynamicButtons.append(restartButton)
        restartButton.grid(row=restartRow, column=5)
        restartRow += 1
        

    stopSelectedButton = ttk.Button(buttonFrame, text='Stop Selected', command=lambda: stopStartRestartSelected('stop'))
    stopSelectedButton.grid(row=1, column=1)

    startSelectedButton = ttk.Button(buttonFrame, text='Start Selected', command=lambda: stopStartRestartSelected('start'))
    startSelectedButton.grid(row=1, column=2)

    restartSelectedButton = ttk.Button(buttonFrame, text='Restart Selected', command=lambda: stopStartRestartSelected('restart'))
    restartSelectedButton.grid(row=1, column=3)

    refreshButton = ttk.Button(buttonFrame, text='Refresh Status', command=updateLabels)
    refreshButton.grid(row=1, column=4)

class ProgressBar:
    def __init__(self, frame):
        self.pb = ttk.Progressbar(frame, mode='indeterminate', length='400')
        self.pb.grid(row=2, column=1)

def prismServiceStatus():
    prismMQStatus = getServiceStatus('PrismMQService')
    if prismMQStatus == 'Running':
        global prismMQLabel
        prismMQLabel = ttk.Label(allFrame, text=prismMQStatus, foreground='green')
        prismMQLabel.grid(row=2, column=2)
    elif prismMQStatus == 'Stopped':
        prismMQLabel = ttk.Label(allFrame, text=prismMQStatus, foreground='red')
        prismMQLabel.grid(row=2, column=2)
    elif prismMQStatus == 'Starting':
        prismMQLabel = ttk.Label(allFrame, text=prismMQStatus, foreground='orange')
        prismMQLabel.grid(row=2, column=2)
    elif prismMQStatus == 'Stopping':
        prismMQLabel = ttk.Label(allFrame, text=prismMQStatus, foreground='orange')
        prismMQLabel.grid(row=2, column=2)
    else:
        prismMQLabel = ttk.Label(allFrame, text=prismMQStatus)
        prismMQLabel.grid(row=2, column=2)

    commonStatus = getServiceStatus('PrismCommonService')
    if commonStatus == 'Running':
        global commonLabel
        commonLabel = ttk.Label(allFrame, text=commonStatus, foreground='green')
        commonLabel.grid(row=3, column=2)
    elif commonStatus == 'Stopped':
        commonLabel = ttk.Label(allFrame, text=commonStatus, foreground='red')
        commonLabel.grid(row=3, column=2)
    elif commonStatus == 'Starting':
        commonLabel = ttk.Label(allFrame, text=commonStatus, foreground='orange')
        commonLabel.grid(row=2, column=2)
    elif commonStatus == 'Stopping':
        commonLabel = ttk.Label(allFrame, text=commonStatus, foreground='orange')
        commonLabel.grid(row=2, column=2)
    else:
        commonLabel = ttk.Label(allFrame, text=commonStatus)
        commonLabel.grid(row=3, column=2)
    
    resiliencyStatus = getServiceStatus('PrismResiliencyService')
    if resiliencyStatus == 'Running':
        global resiliencyLabel
        resiliencyLabel = ttk.Label(allFrame, text=resiliencyStatus, foreground='green')
        resiliencyLabel.grid(row=4, column=2)
    elif resiliencyStatus == 'Stopped':
        resiliencyLabel = ttk.Label(allFrame, text=resiliencyStatus, foreground='red')
        resiliencyLabel.grid(row=4, column=2)
    elif resiliencyStatus == 'Starting':
        resiliencyLabel = ttk.Label(allFrame, text=resiliencyStatus, foreground='orange')
        resiliencyLabel.grid(row=2, column=2)
    elif resiliencyStatus == 'Stopping':
        resiliencyLabel = ttk.Label(allFrame, text=resiliencyStatus, foreground='orange')
        resiliencyLabel.grid(row=2, column=2)
    else:
        resiliencyLabel = ttk.Label(allFrame, text=resiliencyStatus)
        resiliencyLabel.grid(row=4, column=2)
    
    backofficeStatus = getServiceStatus('PrismBackofficeService')
    if backofficeStatus == 'Running':
        global backofficeLabel
        backofficeLabel = ttk.Label(allFrame, text=backofficeStatus, foreground='green')
        backofficeLabel.grid(row=5, column=2)
    elif backofficeStatus == 'Stopped':
        backofficeLabel = ttk.Label(allFrame, text=backofficeStatus, foreground='red')
        backofficeLabel.grid(row=5, column=2)
    elif backofficeStatus == 'Starting':
        backofficeLabel = ttk.Label(allFrame, text=backofficeStatus, foreground='orange')
        backofficeLabel.grid(row=2, column=2)
    elif backofficeStatus == 'Stopping':
        backofficeLabel = ttk.Label(allFrame, text=backofficeStatus, foreground='orange')
        backofficeLabel.grid(row=2, column=2)
    else:
        backofficeLabel = ttk.Label(allFrame, text=backofficeStatus)
        backofficeLabel.grid(row=5, column=2)
    
    V9Status = getServiceStatus('PrismV9Service')
    if V9Status == 'Running':
        global V9Label
        V9Label = ttk.Label(allFrame, text=V9Status, foreground='green')
        V9Label.grid(row=6, column=2)
    elif V9Status == 'Stopped':
        V9Label = ttk.Label(allFrame, text=V9Status, foreground='red')
        V9Label.grid(row=6, column=2)
    elif V9Status == 'Starting':
        V9Label = ttk.Label(allFrame, text=V9Status, foreground='orange')
        V9Label.grid(row=2, column=2)
    elif V9Status == 'Stopping':
        V9Label = ttk.Label(allFrame, text=V9Status, foreground='orange')
        V9Label.grid(row=2, column=2)
    else:
        V9Label = ttk.Label(allFrame, text=V9Status)
        V9Label.grid(row=6, column=2)
    
    licenseStatus = getServiceStatus('PrismLicSvr')
    if licenseStatus == 'Running':
        global licenseLabel
        licenseLabel = ttk.Label(allFrame, text=licenseStatus, foreground='green')
        licenseLabel.grid(row=7, column=2)
    elif licenseStatus == 'Stopped':
        licenseLabel = ttk.Label(allFrame, text=licenseStatus, foreground='red')
        licenseLabel.grid(row=7, column=2)
    elif licenseStatus == 'Starting':
        licenseLabel = ttk.Label(allFrame, text=licenseStatus, foreground='orange')
        licenseLabel.grid(row=2, column=2)
    elif licenseStatus == 'Stopping':
        licenseLabel = ttk.Label(allFrame, text=licenseStatus, foreground='orange')
        licenseLabel.grid(row=2, column=2)
    else:
        licenseLabel = ttk.Label(allFrame, text=licenseStatus)
        licenseLabel.grid(row=7, column=2)
    
    rabbitmqStatus = getServiceStatus('RabbitMQ')
    if rabbitmqStatus == 'Running':
        global rabbitmqLabel
        rabbitmqLabel = ttk.Label(allFrame, text=rabbitmqStatus, foreground='green')
        rabbitmqLabel.grid(row=8, column=2)
    elif rabbitmqStatus == 'Stopped':
        rabbitmqLabel = ttk.Label(allFrame, text=rabbitmqStatus, foreground='red')
        rabbitmqLabel.grid(row=8, column=2)
    elif rabbitmqStatus == 'Starting':
        rabbitmqLabel = ttk.Label(allFrame, text=rabbitmqStatus, foreground='orange')
        rabbitmqLabel.grid(row=2, column=2)
    elif rabbitmqStatus == 'Stopping':
        rabbitmqLabel = ttk.Label(allFrame, text=rabbitmqStatus, foreground='orange')
        rabbitmqLabel.grid(row=2, column=2)
    else:
        rabbitmqLabel = ttk.Label(allFrame, text=rabbitmqStatus)
        rabbitmqLabel.grid(row=8, column=2)
    
    apacheStatus = getServiceStatus('Apache')
    if apacheStatus == 'Running':
        global apacheLabel
        apacheLabel = ttk.Label(allFrame, text=apacheStatus, foreground='green')
        apacheLabel.grid(row=9, column=2)
    elif apacheStatus == 'Stopped':
        apacheLabel = ttk.Label(allFrame, text=apacheStatus, foreground='red')
        apacheLabel.grid(row=9, column=2)
    elif apacheStatus == 'Starting':
        apacheLabel = ttk.Label(allFrame, text=apacheStatus, foreground='orange')
        apacheLabel.grid(row=2, column=2)
    elif apacheStatus == 'Stopping':
        apacheLabel = ttk.Label(allFrame, text=apacheStatus, foreground='orange')
        apacheLabel.grid(row=2, column=2)
    else:
        apacheLabel = ttk.Label(allFrame, text=apacheStatus)
        apacheLabel.grid(row=9, column=2)
    
    mysqlStatus = getServiceStatus('MySQL')
    if mysqlStatus == 'Running':
        global mysqlLabel
        mysqlLabel = ttk.Label(allFrame, text=mysqlStatus, foreground='green')
        mysqlLabel.grid(row=10, column=2)
    elif mysqlStatus == 'Stopped':
        mysqlLabel = ttk.Label(allFrame, text=mysqlStatus, foreground='red')
        mysqlLabel.grid(row=10, column=2)
    elif mysqlStatus == 'Starting':
        mysqlLabel = ttk.Label(allFrame, text=mysqlStatus, foreground='orange')
        mysqlLabel.grid(row=2, column=2)
    elif mysqlStatus == 'Stopping':
        mysqlLabel = ttk.Label(allFrame, text=mysqlStatus, foreground='orange')
        mysqlLabel.grid(row=2, column=2)
    else:
        mysqlLabel = ttk.Label(allFrame, text=mysqlStatus)
        mysqlLabel.grid(row=10, column=2)

    licStatus = getServiceStatus('RProLicSvr')
    if licStatus == 'Running':
        global licLabel
        licLabel = ttk.Label(allFrame, text=licStatus, foreground='green')
        licLabel.grid(row=11, column=2)
    elif licStatus == 'Stopped':
        licLabel = ttk.Label(allFrame, text=licStatus, foreground='red')
        licLabel.grid(row=11, column=2)
    elif licStatus == 'Starting':
        licLabel = ttk.Label(allFrame, text=lictatus, foreground='orange')
        licLabel.grid(row=2, column=2)
    elif licStatus == 'Stopping':
        licLabel = ttk.Label(allFrame, text=licStatus, foreground='orange')
        licLabel.grid(row=2, column=2)
    else:
        licLabel = ttk.Label(allFrame, text=licStatus)
        licLabel.grid(row=11, column=2)

    listnerStatus = getServiceStatus('OracleODS12cr1TNSListener')
    if listnerStatus == 'Running':
        global listnerLabel
        listnerLabel = ttk.Label(allFrame, text=listnerStatus, foreground='green')
        listnerLabel.grid(row=12, column=2)
    elif listnerStatus == 'Stopped':
        listnerLabel = ttk.Label(allFrame, text=listnerStatus, foreground='red')
        listnerLabel.grid(row=12, column=2)
    elif listnerStatus == 'Starting':
        listnerLabel = ttk.Label(allFrame, text=listnerStatus, foreground='orange')
        listnerLabel.grid(row=2, column=2)
    elif listnerStatus == 'Stopping':
        listnerLabel = ttk.Label(allFrame, text=listnerStatus, foreground='orange')
        listnerLabel.grid(row=2, column=2)
    else:
        listnerLabel = ttk.Label(allFrame, text=listnerStatus)
        listnerLabel.grid(row=12, column=2)

    oracleStatus = getServiceStatus('OracleServiceRproODS')
    if oracleStatus == 'Running':
        global oracleLabel
        oracleLabel = ttk.Label(allFrame, text=oracleStatus, foreground='green')
        oracleLabel.grid(row=13, column=2)
    elif oracleStatus == 'Stopped':
        oracleLabel = ttk.Label(allFrame, text=oracleStatus, foreground='red')
        oracleLabel.grid(row=13, column=2)
    elif oracleStatus == 'Starting':
        oracleLabel = ttk.Label(allFrame, text=oracleStatus, foreground='orange')
        oracleLabel.grid(row=2, column=2)
    elif oracleStatus == 'Stopping':
        oracleLabel = ttk.Label(allFrame, text=oracleStatus, foreground='orange')
        oracleLabel.grid(row=2, column=2)
    else:
        oracleLabel = ttk.Label(allFrame, text=oracleStatus)
        oracleLabel.grid(row=13, column=2)

def makeLabels():
    row = 2
    for service in services:
        serviceStatus = getServiceStatus(service)
        if serviceStatus == 'Running':
            serviceLabel = ttk.Label(allFrame, text=serviceStatus, foreground='green')
            dynamicLabels.append(serviceLabel)
            serviceLabel.grid(row=row, column=2)
        elif serviceStatus == 'Stopped':
            serviceLabel = ttk.Label(allFrame, text=serviceStatus, foreground='red')
            dynamicLabels.append(serviceLabel)
            serviceLabel.grid(row=row, column=2)
        else:
            serviceLabel = ttk.Label(allFrame, text=serviceStatus)
            dynamicLabels.append(serviceLabel)
            serviceLabel.grid(row=row, column=2)
        row += 1

def prismServiceStatusDestroy():
    prismMQLabel.destroy()
    commonLabel.destroy()
    resiliencyLabel.destroy()
    backofficeLabel.destroy()
    V9Label.destroy()
    licenseLabel.destroy()
    rabbitmqLabel.destroy()
    apacheLabel.destroy()
    mysqlLabel.destroy()
    oracleLabel.destroy()
    listnerLabel.destroy()
    licLabel.destroy()

def updateLabels():
    prismMQStatus = getServiceStatus('PrismMQService')
    if prismMQStatus == 'Running':
        prismMQLabel.config(text=prismMQStatus, foreground='green')
    elif prismMQStatus == 'Stopped':
        prismMQLabel.config(text=prismMQStatus, foreground='red')
    elif prismMQStatus == 'Starting':
        prismMQLabel.config(text=prismMQStatus, foreground='orange')
    elif prismMQStatus == 'Stopping':
        prismMQLabel.config(text=prismMQStatus, foreground='orange')
    else:
        prismMQLabel.config(text=prismMQStatus)
        
    commonStatus = getServiceStatus('PrismCommonService')
    if commonStatus == 'Running':
        commonLabel.config(text=commonStatus, foreground='green')
    elif commonStatus == 'Stopped':
        commonLabel.config(text=commonStatus, foreground='red')
    elif commonStatus == 'Starting':
        commonLabel.config(text=commonStatus, foreground='orange')
    elif commonStatus == 'Stopping':
        commonLabel.config(text=commonStatus, foreground='orange')
    else:
        commonLabel.config(text=commonStatus)
    
    resiliencyStatus = getServiceStatus('PrismResiliencyService')
    if resiliencyStatus == 'Running':
        resiliencyLabel.config(text=resiliencyStatus, foreground='green')
    elif resiliencyStatus == 'Stopped':
        resiliencyLabel.config(text=resiliencyStatus, foreground='red')
    elif resiliencyStatus == 'Starting':
        resiliencyLabel.config(text=resiliencyStatus, foreground='orange')
    elif resiliencyStatus == 'Stopping':
        resiliencyLabel.config(text=resiliencyStatus, foreground='orange')
    else:
        resiliencyLabel.config(text=resiliencyStatus)
    
    backofficeStatus = getServiceStatus('PrismBackofficeService')
    if backofficeStatus == 'Running':
        backofficeLabel.config(text=backofficeStatus, foreground='green')
    elif backofficeStatus == 'Stopped':
        backofficeLabel.config(text=backofficeStatus, foreground='red')
    elif backofficeStatus == 'Starting':
        backofficeLabel.config(text=backofficeStatus, foreground='orange')
    elif backofficeStatus == 'Stopping':
        backofficeLabel.config(text=backofficeStatus, foreground='orange')
    else:
        backofficeLabel.config(text=backofficeStatus)
    
    V9Status = getServiceStatus('PrismV9Service')
    if V9Status == 'Running':
        V9Label.config(text=V9Status, foreground='green')
    elif V9Status == 'Stopped':
        V9Label.config(text=V9Status, foreground='red')
    elif V9Status == 'Starting':
        V9Label.config(text=V9Status, foreground='orange')
    elif V9Status == 'Stopping':
        V9Label.config(text=V9Status, foreground='orange')
    else:
        V9Label.config(text=V9Status)
    
    licenseStatus = getServiceStatus('PrismLicSvr')
    if licenseStatus == 'Running':
        licenseLabel.config(text=licenseStatus, foreground='green')
    elif licenseStatus == 'Stopped':
        licenseLabel.config(text=licenseStatus, foreground='red')
    elif licenseStatus == 'Starting':
        licenseLabel.config(text=licenseStatus, foreground='orange')
    elif licenseStatus == 'Stopping':
        licenseLabel.config(text=licenseStatus, foreground='orange')
    else:
        licenseLabel.config(text=licenseStatus)
    
    rabbitmqStatus = getServiceStatus('RabbitMQ')
    if rabbitmqStatus == 'Running':
        rabbitmqLabel.config(text=rabbitmqStatus, foreground='green')
    elif rabbitmqStatus == 'Stopped':
        rabbitmqLabel.config(text=rabbitmqStatus, foreground='red')
    elif rabbitmqStatus == 'Starting':
        rabbitmqLabel.config(text=rabbitmqStatus, foreground='orange')
    elif rabbitmqStatus == 'Stopping':
        rabbitmqLabel.config(text=rabbitmqStatus, foreground='orange')
    else:
        rabbitmqLabel.config(text=rabbitmqStatus)
    
    apacheStatus = getServiceStatus('Apache')
    if apacheStatus == 'Running':
        apacheLabel.config(text=apacheStatus, foreground='green')
    elif apacheStatus == 'Stopped':
        apacheLabel.config(text=apacheStatus, foreground='red')
    elif apacheStatus == 'Starting':
        apacheLabel.config(text=apacheStatus, foreground='orange')
    elif apacheStatus == 'Stopping':
        apacheLabel.config(text=apacheStatus, foreground='orange')
    else:
        apacheLabel.config(text=apacheStatus)
    
    mysqlStatus = getServiceStatus('MySQL')
    if mysqlStatus == 'Running':
        mysqlLabel.config(text=mysqlStatus, foreground='green')
    elif mysqlStatus == 'Stopped':
        mysqlLabel.config(text=mysqlStatus, foreground='red')
    elif mysqlStatus == 'Starting':
        mysqlLabel.config(text=mysqlStatus, foreground='orange')
    elif mysqlStatus == 'Stopping':
        mysqlLabel.config(text=mysqlStatus, foreground='orange')
    else:
        mysqlLabel.config(text=mysqlStatus)

    licStatus = getServiceStatus('RProLicSvr')
    if licStatus == 'Running':
        licLabel.config(text=licStatus, foreground='green')
    elif licStatus == 'Stopped':
        licLabel.config(text=licStatus, foreground='red')
    elif licStatus == 'Starting':
        licLabel.config(text=licStatus, foreground='orange')
    elif licStatus == 'Stopping':
        licLabel.config(text=licStatus, foreground='orange')
    else:
        licLabel.config(text=licStatus)

    oracleStatus = getServiceStatus('OracleServiceRproODS')
    if oracleStatus == 'Running':
        oracleLabel.config(text=oracleStatus, foreground='green')
    elif oracleStatus == 'Stopped':
        oracleLabel.config(text=oracleStatus, foreground='red')
    elif oracleStatus == 'Starting':
        oracleLabel.config(text=oracleStatus, foreground='orange')
    elif oracleStatus == 'Stopping':
        oracleLabel.config(text=oracleStatus, foreground='orange')
    else:
        oracleLabel.config(text=oracleStatus)

    listnerStatus = getServiceStatus('OracleODS12cr1TNSListener')
    if listnerStatus == 'Running':
        listnerLabel.config(text=listnerStatus, foreground='green')
    elif listnerStatus == 'Stopped':
        listnerLabel.config(text=listnerStatus, foreground='red')
    elif listnerStatus == 'Starting':
        listnerLabel.config(text=listnerStatus, foreground='orange')
    elif listnerStatus == 'Stopping':
        listnerLabel.config(text=listnerStatus, foreground='orange')
    else:
        listnerLabel.config(text=listnerStatus)

def makeCheckBoxes():
    row = 2

    global selectCheckBox
    selectCheckBox = ttk.Checkbutton(allFrame, text='Select/Deselect All', command=selectDeselect)
    selectCheckBox.grid(row=1, column=1, sticky='w')
    selectCheckBox.invoke() 
    
    for service in services:
        serviceCheckBox = ttk.Checkbutton(allFrame, text=service)
        serviceCheckBox.grid(row=row, column=1, sticky='w')
        dynamicCheckBoxes.append(serviceCheckBox)
        row += 1

    for serviceCheckBox in dynamicCheckBoxes:
        serviceCheckBox.invoke()

def selectDeselect():
    if selectCheckBox.instate(['selected']):
        for dynamicCheckBox in dynamicCheckBoxes:
            dynamicCheckBox.state(['selected'])

    if selectCheckBox.instate(['!selected']):
        for dynamicCheckBox in dynamicCheckBoxes:
            dynamicCheckBox.state(['!selected'])

def getServiceStatus(serviceName):
    try:
        service = win32serviceutil.QueryServiceStatus(serviceName)[1]
    except:
        return 'Not Installed'
    if service == 4:
        return 'Running'
    if service == 1:
        return 'Stopped'
    if service == 3:
        return 'Stopping'
    if service == 2:
        return 'Starting'

def stopStartRestart(service, action):
    def stopPrismMQService():
        if getServiceStatus('PrismMQService') == 'Running':
            pb.pb.start()
            try:
                win32serviceutil.StopService('PrismMQService')
            except Exception as ex:
                messagebox.showerror('Error', ex)
                pb.pb.stop()
                return
            checkStatus('PrismMQService', 'stop')
            updateLabels()
    def stopService():
        if getServiceStatus(service) == 'Running':
            pb.pb.start()
            try:
                win32serviceutil.StopService(service)
            except Exception as ex:
                messagebox.showerror('Error', ex)
                pb.pb.stop()
                print('except')
                return
            checkStatus(service, 'stop')
            updateLabels()
            pb.pb.stop()
    def startRabbitMQService():
        if getServiceStatus('RabbitMQ') == 'Stopped':
            pb.pb.start()
            try:
                win32serviceutil.StartService('RabbitMQ')
            except Exception as ex:
                messagebox.showerror('Error', ex)
                pb.pb.stop()
                return
            checkStatus('RabbitMQ', 'start')
            updateLabels()
    def startService():
        if getServiceStatus(service) == 'Stopped':
            pb.pb.start()
            try:
                win32serviceutil.StartService(service)
            except Exception as ex:
                messagebox.showerror('Error', ex)
                pb.pb.stop()
                return
            checkStatus(service, 'start')
            updateLabels()
            pb.pb.stop()
    def startPrismMQService():
        if getServiceStatus('PrismMQService') == 'Stopped':
            pb.pb.start()
            try:
                win32serviceutil.StartService('PrismMQService')
            except Exception as ex:
                messagebox.showerror('Error', ex)
                pb.pb.stop()
                return
            checkStatus('PrismMQService', 'start')
            updateLabels()
            pb.pb.stop()

    if action == 'stop':
        if service == 'RabbitMQ':   
            stopPrismMQService()
            stopService()
        else:     
            stopService()

    if action == 'start':
        if service == 'PrismMQService':
            startRabbitMQService()
            startService()
            
        else:
            startService()

    if action == 'restart':
            
        # Stop Services
        if service == 'RabbitMQ':
            stopPrismMQService()
            stopService()
        else:
            stopService()

        # Start Services
        if service == 'PrismMQService':
            startRabbitMQService()
            startService()

        if service == 'RabbitMQ':
            startService()
            startPrismMQService()
        else:
            startService()
            
def stopStartRestartSelected(action):
    def stopPrismMQService():
        pb.pb.start()
        try:
            win32serviceutil.StopService('PrismMQService')
        except Exception as ex:
            messagebox.showerror('Error', ex)
            pb.pb.stop()
            return
        checkStatus('PrismMQService', 'stop')
        updateLabels()
        pb.pb.stop()
    def stopService():
        pb.pb.start()
        try:
            win32serviceutil.StopService(dynamicCheckBox['text'])
        except Exception as ex:
            messagebox.showerror('Error', ex)
            pb.pb.stop()
            return
        checkStatus(dynamicCheckBox['text'], 'stop')
        updateLabels()
        pb.pb.stop()
    def startRabbitMQService():
        pb.pb.start()
        try:
            win32serviceutil.StartService('RabbitMQ')
        except Exception as ex:
            messagebox.showerror('Error', ex)
            pb.pb.stop()
            return
        checkStatus('RabbitMQ', 'start')
        updateLabels()
        pb.pb.stop()
    def startService():
        pb.pb.start()
        try:
            win32serviceutil.StartService(dynamicCheckBox['text'])
        except Exception as ex:
            messagebox.showerror('Error', ex)
            pb.pb.stop()
            return
        checkStatus(dynamicCheckBox['text'], 'start')
        updateLabels()
        pb.pb.stop()
    def startPrismMQService():
        pb.pb.start()
        try:
            win32serviceutil.StartService('PrismMQService')
        except Exception as ex:
            messagebox.showerror('Error', ex)
            pb.pb.stop()
            return
        checkStatus('PrismMQService', 'start')
        updateLabels()
        pb.pb.stop()
    if action == 'stop':
        for dynamicCheckBox in dynamicCheckBoxes:
            if dynamicCheckBox.instate(['selected']):
                if dynamicCheckBox['text'] == 'RabbitMQ':
                    if getServiceStatus('PrismMQService') == 'Running':
                        stopPrismMQService()
                
                    if getServiceStatus(dynamicCheckBox['text']) == 'Running':
                        stopService()
                    
                else:
                    if getServiceStatus(dynamicCheckBox['text']) == 'Running':
                        stopService()

    if action == 'start':
        for dynamicCheckBox in reversed(dynamicCheckBoxes):
            if dynamicCheckBox.instate(['selected']):
                if dynamicCheckBox['text'] == 'PrismMQService':
                    if getServiceStatus('Rabbit') == 'Stopped':
                        startRabbitMQService()
                    
                    if getServiceStatus(dynamicCheckBox['text']) == 'Stopped':
                        startService()
                    
                else:
                    if getServiceStatus(dynamicCheckBox['text']) == 'Stopped':
                        startService()
                    
    if action == 'restart':
        # Stop Services
        for dynamicCheckBox in dynamicCheckBoxes:
            if dynamicCheckBox.instate(['selected']):
                if dynamicCheckBox['text'] == 'RabbitMQ':
                    if getServiceStatus('PrismMQService') == 'Running':
                        stopPrismMQService()
                    if getServiceStatus(dynamicCheckBox['text']) == 'Running':
                        stopService()
                    
                else:
                    if getServiceStatus(dynamicCheckBox['text']) == 'Running':
                        stopService()
                        updateLabels()
                    
        # Start Services
        for dynamicCheckBox in reversed(dynamicCheckBoxes):
            if dynamicCheckBox.instate(['selected']):
                if dynamicCheckBox['text'] == 'PrismMQService':
                    if getServiceStatus('RabbitMQ') == 'Stopped':
                        startRabbitMQService()
                    
                    if getServiceStatus(dynamicCheckBox['text']) == 'Stopped':
                        startService()
                    
                if dynamicCheckBox['text'] == 'RabbitMQ':
                    if getServiceStatus('PrismMQService') == 'Stopped':
                        
                        startPrismMQService()
                    
                    if getServiceStatus(dynamicCheckBox['text']) == 'Stopped':
                        startService()
                    
                else:
                    if getServiceStatus(dynamicCheckBox['text']) == 'Stopped':
                        startService()

def checkStatus(service, action):
    if action == 'stop':
        currentStatus = getServiceStatus(service)
        while currentStatus != 'Stopped':
            updateLabels()
            mainWindow.update()
            currentStatus = getServiceStatus(service)
        
        if currentStatus == 'Stopped':
            updateLabels()
            mainWindow.update()
        '''    
        if currentStatus != 'Stopped':
            currentStatus = getServiceStatus(service)
            updateLabels()
            mainWindow.update()
            checkStatus(service, action)
        '''

    if action == 'start':
        currentStatus = getServiceStatus(service)
        while currentStatus != 'Running':
            updateLabels()
            mainWindow.update()
            currentStatus = getServiceStatus(service)

        if currentStatus == 'Running':
            updateLabels()
            mainWindow.update()
        '''     
        if currentStatus != 'Running':
            currentStatus = getServiceStatus(service)
            updateLabels()
            mainWindow.update()
            checkStatus(service, action)
        '''
     
if __name__ == '__main__':

    mainWindow = tk.Tk()
    mainWindow.title('Rpro Service Manager')

    allFrame = tk.Frame(mainWindow)
    allFrame.pack(padx=20, pady=20)

    buttonFrame = tk.Frame(allFrame)
    buttonFrame.grid(row=17, column=1, columnspan=5, pady=(20, 0))

    pbFrame = tk.Frame(allFrame)
    pbFrame.grid(row=16, column=1, columnspan=5, pady=(20, 0))

    makeCheckBoxes()
    prismServiceStatus()
    makeButtons()

    pb = ProgressBar(pbFrame)

    print(dynamicButtons)
    print(dynamicLabels)

    
    mainWindow.mainloop()
