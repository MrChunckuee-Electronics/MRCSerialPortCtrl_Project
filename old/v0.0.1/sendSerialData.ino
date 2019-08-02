/*******************************************************************************
 *
 *    Envio de datos al puerto serie
 *
 *******************************************************************************
 * FileName:        sendSerialData
 * Processor:       EX-LM4F120XL
 * Complier:        Energia 0101e0017
 * Author:          Pedro Sánchez (MrChunckuee)
 * Blog:            http://mrchunckuee.blogspot.com/
 * Email:           mrchunckuee.psr@gmail.com
 * Description:     Este ejemplo envia un valor tipo Int atraves del puerto serie, 
 *                  para ser procesado en Python
 *******************************************************************************
 * Rev.         Date            Comment
 *   v0.0.0     22/02/2019      Creación del firmware
 ******************************************************************************/

unsigned int send_data;

void setup() {
  Serial.begin(9600);
  pinMode(RED_LED, OUTPUT);     
  pinMode(GREEN_LED, OUTPUT); 
  pinMode(BLUE_LED, OUTPUT); 
  pinMode(PUSH1, INPUT_PULLUP);
  //Clear all
  send_data = 0;
  digitalWrite(RED_LED, LOW);
  digitalWrite(GREEN_LED, LOW);
  digitalWrite(BLUE_LED, LOW);
}

void loop() {
    Serial.println(send_data);
    send_data ++;
    if(send_data == 400)
      send_data= 0;
    digitalWrite(BLUE_LED, HIGH);
    delay(500);
    digitalWrite(BLUE_LED, LOW);
    delay(500);           
  }
