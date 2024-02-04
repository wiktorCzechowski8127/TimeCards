##########################################################
# Imports
##########################################################
import time
import RPi.GPIO as GPIO
GPIO.setmode(GPIO.BCM)

##########################################################
# Define
##########################################################
BUZZER = 12
YELLOW_LED = 20
GREEN_LED = 21
FAILITURE_TIME = 1.60

##########################################################
# Variables
##########################################################
signalLedPort = GREEN_LED
unusedSignalLedPort = YELLOW_LED

##########################################################
# Functions
##########################################################
def gpioInitialize():
    GPIO.setup(BUZZER, GPIO.OUT)
    GPIO.setup(YELLOW_LED,GPIO.OUT)
    GPIO.setup(GREEN_LED,GPIO.OUT)

    GPIO.output(BUZZER,False)
    GPIO.output(YELLOW_LED,False)
    GPIO.output(GREEN_LED,False)


def blinkSuccess():
    GPIO.output(YELLOW_LED, False)
    GPIO.output(GREEN_LED, False)
    
    GPIO.output(GREEN_LED, True)
    GPIO.output(BUZZER, True)
    time.sleep(0.05)

    GPIO.output(GREEN_LED, False)
    GPIO.output(BUZZER, False)
    time.sleep(0.05)

    GPIO.output(GREEN_LED, True)
    GPIO.output(BUZZER, True)
    time.sleep(0.05)
    
    GPIO.output(GREEN_LED, False)
    GPIO.output(BUZZER, False)


def blinkFailiture():

    GPIO.output(YELLOW_LED, False)
    GPIO.output(GREEN_LED, False)
    GPIO.output(BUZZER, True)

    sec = 0
    ledStatus = True
    while(sec < FAILITURE_TIME):
        GPIO.output(YELLOW_LED, ledStatus)
        time.sleep(0.10)
        ledStatus = not(ledStatus)
        sec += 0.10

    GPIO.output(BUZZER, False)


def readingTokenLed():
    GPIO.output(YELLOW_LED, True)
    GPIO.output(GREEN_LED, False)


def blink(now, lastSecond, isSafeMode, signalLed):
    
    if(now.second != lastSecond):
        lastSecond = now.second

        if (isSafeMode):
            signalLedPort = YELLOW_LED
            unusedSignalLedPort = GREEN_LED
        else:
            signalLedPort = GREEN_LED
            unusedSignalLedPort = YELLOW_LED

        signalLed = not(signalLed)
        GPIO.output(signalLedPort, signalLed)
        GPIO.output(unusedSignalLedPort, False)

    return signalLed, lastSecond

def excelProcessed():
    GPIO.output(GREEN_LED, True)
    GPIO.output(YELLOW_LED, True)