/*******************************************************************************
 *
 *          ESP32: Comunicacion serie con Python
 *
 *******************************************************************************
 * FileName:        ESP32-S2-Saola-1R_GUIPython_SendCounterData.ino
 * Processor:       ESP32-S2-WROVER
 * Complier:        Arduino IDE v1.8.19
 * Blog:            http://mrchunckuee.blogspot.com/
 * Description:     Enviar datos al puerto serial, espera un mensaje de "INICIAR" 
 *                  desde el puerto serial y luego envía una cadena de texto y 
 *                  un contador cada segundo.
 *******************************************************************************
 * Rev.         Date            Comment
 *   v1.0.0     13/19/2025      - Creación y prueba de funcionamiento el con  
 *                                script de Python "MRC SerialPortCtrl".
 ******************************************************************************/
#include <Adafruit_NeoPixel.h>
 
// ESP32S2 SAOLA 1R GPIO the WS2812 Led.
#define PIN        18 

//Single WS2812 config
Adafruit_NeoPixel pixels(1, PIN, NEO_GRB + NEO_KHZ800);

#define BAUDRATE 115200

const long interval = 1000;
static unsigned long previousMillis = 0;
static int counter = 0;
bool canSend = false;

void setup() {
  pixels.setBrightness(50);
  pixels.begin(); 
  pixels.setPixelColor(0, 0xFF 0000);
  pixels.show();  // WS2812 Led = Red
  Serial.begin(BAUDRATE);
  Serial.println("LISTO..."); 
}

void loop() {
  // Espera a recibir INICIAR
  if (!canSend) {
    if (Serial.available() > 0) {
      String receivedData = Serial.readStringUntil('\n');
      receivedData.trim(); 
      
      if (receivedData.equalsIgnoreCase("INICIAR")) {
        pixels.setPixelColor(0, 0x00FF00);
        pixels.show();  // WS2812 Led = Green
        canSend = true;
        Serial.println("Comenzando la transmision...");
      }
    }
    return; 
  }

  unsigned long currentMillis = millis();
  if (currentMillis - previousMillis >= interval) {
    previousMillis = currentMillis;
    Serial.print("Counter: ");
    Serial.println(counter);
    counter++;
  }
}
