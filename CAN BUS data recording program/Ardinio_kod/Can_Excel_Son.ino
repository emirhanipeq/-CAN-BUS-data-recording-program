  #include <SPI.h>
  #include "mcp_can.h"

  const int SPI_CS_PIN = 10;
  MCP_CAN CAN(SPI_CS_PIN);

  void setup()
  {
    Serial.begin(115200);

    if (CAN.begin(MCP_ANY, CAN_500KBPS, MCP_8MHZ) == CAN_OK)
      Serial.println("CAN BUS Başlatildi");
    else {
      Serial.println("CAN BUS Başlatilamadi");
      while (1);
    }

    CAN.setMode(MCP_NORMAL);  // Normal mod
  }

  void loop()
  {
    long unsigned int rxId;
    unsigned char len = 0;
    unsigned char rxBuf[8];

    if (CAN_MSGAVAIL == CAN.checkReceive())
    {
      CAN.readMsgBuf(&rxId, &len, rxBuf);

      // C# uyumlu seri formatta gönder
      Serial.print("ID: ");
      Serial.print(rxId, HEX);  // C# kodunda bu alınacak
      Serial.print(" DATA: ");
      for (int i = 0; i < len; i++) {
        if (rxBuf[i] < 0x10) Serial.print("0"); // Tek haneliler için 0 ekle
        Serial.print(rxBuf[i], HEX);
        Serial.print(" ");
      }
      Serial.println();
      delay(1); 
    }
  }
