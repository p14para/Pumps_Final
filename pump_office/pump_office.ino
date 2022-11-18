/**
   ----------------------------------------------------------------------------
   https://github.com/miguelbalboa/rfid
   CARDS: MIFARE Classic PICC 1K

   Typical pin layout used:
                         MEGA     UNO
   RST/Reset   RST          8       8
   SPI SS      SDA(SS)      9       9
   SPI MOSI    MOSI         51     11
   SPI MISO    MISO         50     12
   SPI SCK     SCK          52     13 

         ICSP UNO
   MISO   ------  VCC
           |  |
    SCK   ------  MOSI
           |  |
  RESET   ------  GRD
*/

#include <SPI.h>
#include <MFRC522.h>

#define MFRC522_RST_PIN  8
#define MFRC522_SS_PIN   9

MFRC522 mfrc522(MFRC522_SS_PIN, MFRC522_RST_PIN);
MFRC522::MIFARE_Key key;

String serial_read = "";
String nfc_write_buffer = "";
String nfc_card_uid = "";
int nfc_new_init = 0;

void setup() {
  Serial.begin(9600); // Initialize serial communications with the PC
  SPI.begin();        // Init SPI bus
  mfrc522.PCD_Init(); // Init MFRC522 card

  // Prepare the key (used both as key A and as key B) FFFFFFFFFFFFh factory default
  for (byte i = 0; i < 6; i++) {
    key.keyByte[i] = 0xFF;
  }

  Serial.print(F("Key: "));
  dump_byte_array(key.keyByte, MFRC522::MF_KEY_SIZE);
  Serial.println();
}


// ------------------------------------------------------------------------------
void loop() {
  // Look for new cards

  if (nfc_new_init == 255) {
      nfc_write_buffer = "";
      nfc_card_uid = "";
      nfc_new_init = 0;
      mfrc522.PCD_Reset();
      mfrc522.PCD_Init();
  }
  
  if ( !mfrc522.PICC_IsNewCardPresent()) {
    if (nfc_write_buffer.length() != 16) {
      return;
    }
  } else {
    nfc_new_init = 1;
  }

  if (nfc_write_buffer.length() == 16) {
    if (nfc_new_init == 0) {
      mfrc522.PCD_Reset();
      mfrc522.PCD_Init();
      nfc_new_init = 1;
    }
  }
  
  // Select one of the cards
  if ( !mfrc522.PICC_ReadCardSerial()) {
     return;
  }

  nfc_new_init = 0;

  Serial.print(F("@"));
  
  Serial.print(F("CUID#"));
  dump_byte_array(mfrc522.uid.uidByte, mfrc522.uid.size);
  Serial.println("#");
  
  Serial.print(F("PICC#"));
  MFRC522::PICC_Type piccType = mfrc522.PICC_GetType(mfrc522.uid.sak);
  Serial.print(mfrc522.PICC_GetTypeName(piccType));
  Serial.println("#");

  Serial.print(F("NFCWB#"));
  Serial.print(nfc_write_buffer);
  Serial.println("#");
  
  // Check for compatibility
  if (piccType != MFRC522::PICC_TYPE_MIFARE_MINI &&  piccType != MFRC522::PICC_TYPE_MIFARE_1K &&  piccType != MFRC522::PICC_TYPE_MIFARE_4K) {
    Serial.println(F("Unknown NFC type."));
    return;
  }

  byte sector         = 1;
  byte blockAddr      = 4;
  byte trailerBlock   = 7;
  MFRC522::StatusCode status;
  byte buffer[18];
  byte size = sizeof(buffer);

  // Read Authenticate using keyA
  status = (MFRC522::StatusCode) mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A, trailerBlock, &key, &(mfrc522.uid));
  if (status != MFRC522::STATUS_OK) {
    Serial.print(F("AUTHFA:"));
    Serial.println(mfrc522.GetStatusCodeName(status));
    return;
  }

  // Read data from the block
  status = (MFRC522::StatusCode) mfrc522.MIFARE_Read(blockAddr, buffer, &size);
  if (status != MFRC522::STATUS_OK) {
    Serial.print(F("Read failed. "));
    Serial.println(mfrc522.GetStatusCodeName(status));
  }
  
  String bstr = (F("DATA#"));
  String data_block = (char*)buffer;
  bstr += data_block;
  bstr += "#";
  Serial.println(bstr);


  // Write to NFC
  if (nfc_write_buffer.length()) {
    status = (MFRC522::StatusCode) mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_B, trailerBlock, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
      Serial.print(F("AUTHFB:"));
      Serial.println(mfrc522.GetStatusCodeName(status));
      return;
    }

    byte astr[nfc_write_buffer.length() + 1];
    nfc_write_buffer.getBytes(astr, sizeof(astr));

    Serial.println(string_byte_array(mfrc522.uid.uidByte, mfrc522.uid.size));
    Serial.println(nfc_card_uid);

    if (string_byte_array(mfrc522.uid.uidByte, mfrc522.uid.size) == nfc_card_uid) {
      Serial.println("Identical nfc_uid");
    } else {
      nfc_write_buffer = "";
      nfc_card_uid = "";
      nfc_new_init = 255;
      Serial.println("#CHFAEND#");
      return;
    }

    // Write data to the block
    Serial.print(F("Writing data into block "));
    Serial.println(blockAddr);
    dump_byte_array(astr, 16);
    Serial.println();

    Serial.println(F("WRITE"));
    status = (MFRC522::StatusCode) mfrc522.MIFARE_Write(blockAddr, astr, 16);
    if (status != MFRC522::STATUS_OK) {
      Serial.print(F("Write failed. "));
      Serial.println(mfrc522.GetStatusCodeName(status));
    }
    Serial.println();
  
    Serial.println(F("CHECK"));
    status = (MFRC522::StatusCode) mfrc522.MIFARE_Read(blockAddr, buffer, &size);
    if (status != MFRC522::STATUS_OK) {
      Serial.print(F("Check failed. "));
      Serial.println(mfrc522.GetStatusCodeName(status));
    }

    String bstr = (F("DATAW#"));
    String data_block = (char*)buffer;
    bstr += data_block;
    bstr += "#";
    Serial.println(bstr);
    
    Serial.println(F("Checking"));
    byte count = 0;
    for (byte i = 0; i < 16; i++) {
      if (buffer[i] == astr[i])
        count++;
    }
    Serial.print(F("Number of bytes that match: ")); Serial.println(count);
    if (count == 16) {
      Serial.println(F("#CHOK"));
    } else {
      Serial.println(F("#CHFA"));
    }
    nfc_card_uid = "";
    nfc_write_buffer = "";
  }

  // Halt PICC
  mfrc522.PICC_HaltA();
  // Stop encryption on PCD
  mfrc522.PCD_StopCrypto1();
  Serial.println("END#");
}




void serialEvent() {
  
  char inChar = (char)Serial.read();
  
  if (inChar == '\n') {
    Serial.print("Receive "); Serial.println(serial_read);
    
    if (nfc_write_buffer.length()) {
      Serial.println("There is nfc write buffer...");
      serial_read = "";
      return;
    }

    nfc_write_buffer = "";
    
    if (String(serial_read).length() == 20) {
      nfc_write_buffer = serial_read.substring(0, 16);
      nfc_card_uid = serial_read.substring(12);
    }
    
    Serial.print("WBuffer "); Serial.print(nfc_write_buffer); Serial.print(" :"); Serial.println(nfc_write_buffer.length());
    Serial.print("WCardid "); Serial.println(nfc_card_uid); Serial.print(" :"); Serial.println(nfc_card_uid.length());
    Serial.println(String(serial_read).length());
    
    serial_read = "";
  } else {
    serial_read += inChar;
  }
}





void dump_byte_array(byte *buffer, byte bufferSize) {
  for (byte i = 0; i < bufferSize; i++) {
    Serial.print(buffer[i] < 0x10 ? "0" : "");
    Serial.print(buffer[i], HEX);
  }
}

String string_byte_array(byte *buffer, byte bufferSize) {
  String bstr = "";
  for (byte i = 0; i < bufferSize; i++) {
    bstr += String(buffer[i] < 0x10 ? "0" : "");
    bstr += String(buffer[i], HEX);
  }
  bstr.toUpperCase();
  return bstr;
}
