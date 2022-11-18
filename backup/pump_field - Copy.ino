/*
  Hold button 1 five seconds -> LCD card size and free space, Log size, Pump Id
                                Serial Print full SD-Card info
  Hold button 2 five seconds -> DELETE logxxxx.dat
*/

#include <Wire.h>
#include <TimeLib.h>
#include <DS1307RTC.h>
#include <LiquidCrystal.h>
#include <SPI.h>
#include <MFRC522.h>
#include "SdFat.h"

// GLOBAL
String pump_id = "1001";
char log_file[12];

// LCD
LiquidCrystal lcd(36, 34, 32, 30, 28, 26);

// BUTTONS
byte button_1 = 38;
byte button_2 = 40;
byte button_3 = 42;
byte button_1_state = 0;
byte button_2_state = 0;
byte button_3_state = 0;
byte button_1_memstate = 0;
byte button_2_memstate = 0;
byte button_3_memstate = 0;
byte button_1_holdcount = 0;
byte button_2_holdcount = 0;
byte button_3_holdcount = 0;

// MENU
#define menu_idle 0
#define menu_ask_delete_log 1
#define menu_ask_hours 2
#define menu_ask_start 3
#define menu_set_start 4
#define menu_show_remain_work 5
#define menu_set_date 6
#define menu_set_time 7
#define menu_set_name 8
byte menu_state = menu_idle;
int set_hours = 0;
int remain_hours = 0;
int hours_limit = 12;
int idle_time_count = 0;
String pump_work_until = "";
long pump_end_time;


// NFC
#define MFRC522_RST_PIN 8
#define MFRC522_SS_PIN 9
MFRC522 mfrc522(MFRC522_SS_PIN, MFRC522_RST_PIN);
MFRC522::MIFARE_Key key;
String serial_read = "";
String nfc_write_buffer = "";
String nfc_card_uid = "";
byte nfc_new_init = 0;
int fails_count = 0;

// SD
#define USE_SDIO 0
const uint8_t SD_CHIP_SELECT = SS;
const int8_t DISABLE_CHIP_SELECT = -1;
SdFile file;

#if USE_SDIO
// Use faster SdioCardEX
SdFatSdioEX sd;
// SdFatSdio sd;
#else // USE_SDIO
SdFat sd;
#endif  // USE_SDIO
ArduinoOutStream cout(Serial);
uint32_t cardSize;
uint32_t eraseSize;
#define sdErrorMsg(msg) sd.errorPrint(F(msg));

// RTC
int mem_last_second = 0;

void setup() {
  Serial.begin(9600);
  Serial.println(F("Boot..."));

  Serial.print(F("PUMP ID:"));
  Serial.println(pump_id);

  String bstr = "log" + pump_id + ".dat";
  for (int i = 0; i < bstr.length(); i++) {
    log_file[i] = bstr[i];
  }

  Serial.print(F("LOG FILE:"));
  Serial.println(log_file);

  // LCD
  Serial.println(F("LCD initialize."));
  lcd.begin(16, 2);
  lcd.setCursor(0,0);
  lcd.print(F("Initializing..."));
  delay(2000);

  /*
    char j;
    for (int i = 110; i < 150; i++) {
      lcd.clear();
      lcd.setCursor(0, 0);
      lcd.print(i);
      lcd.setCursor(0, 1);
      lcd.print(j = i);
      delay(2000);
    }
  */

  // NFC
  Serial.println(F("NFC initialize."));
  SPI.begin();
  mfrc522.PCD_Init();

  // Prepare the key (used both as key A and as key B) FFFFFFFFFFFFh factory default
  for (byte i = 0; i < 6; i++) {
    key.keyByte[i] = 0xFF;
  }
  Serial.print(F("Key: "));
  dump_byte_array(key.keyByte, MFRC522::MF_KEY_SIZE);
  Serial.println();

  // BUTTONS
  Serial.println(F("Buttons set."));
  pinMode (button_1, INPUT);
  pinMode (button_2, INPUT);
  pinMode (button_3, INPUT);

  // RTC
  Serial.println(F("RTC initialize."));
  while (!Serial) ; // wait for serial
  delay(200);

  // RELAY
  Serial.println(F("Relay set."));
  pinMode (22, OUTPUT);
  digitalWrite(22, HIGH);

  // SD
  Serial.println(F("SD initialize."));
  if (!sd.begin(SD_CHIP_SELECT, SD_SCK_MHZ(50))) {
    Serial.println(F("  SD initialization error."));
  } else {
    file.close();
    Serial.println(F("Check if log_it already exists."));
    if (!sd.exists(log_file)) {
      Serial.println(F("  Not found. Trying to create it."));
      if (!file.open(log_file, O_CREAT | O_WRITE | O_EXCL)) {
        Serial.println(F("  Unable to create file."));
      } else {
        Serial.println(F("  Log file created."));
        log_it(pump_id);   // New file. Write pump id
      }
    } else {
      if (!file.open(log_file, O_WRITE | O_APPEND)) {
        Serial.println(F("  Error opening log_it."));
      } else {
        Serial.print(F("  Log file opened. Size: "));
        Serial.println(String(file.fileSize()));
        log_it(get_time_stamp(false) + " restart");
      }
    }
  }
}

void loop() {
  tmElements_t tm;
  String bstrt;
  String bstrd;
  String bstr;

  if (menu_state != menu_idle && menu_state != menu_show_remain_work) {
    idle_time_count++;
    if (idle_time_count >= 700) {
      menu_state = menu_idle;
      idle_time_count = 0;
    }
  }

  if (menu_state == menu_ask_hours) {
    lcd_shout(F("vPEs KAPTAs"), 0, 0);
    lcd.setCursor(12, 0);
    lcd.print(remain_hours);
    lcd.print("  ");
    lcd_shout(F("pOsEs vPEs"), 0, 1);
    lcd.setCursor(11, 1);
    lcd.print(set_hours);
    lcd.print(" ");
  }

  if (menu_state == menu_ask_start) {
    lcd_shout(F("BAlATE vPEs:"), 0, 0);
    lcd.setCursor(12, 0);
    lcd.print(set_hours);
    lcd_shout(F("+jEKINA   -AKYPO"), 0, 1);
  }

  if (menu_state == menu_show_remain_work) {
    lcd_shout(F("pOTIsMA MEXPI"), 0, 0);
    lcd.setCursor(0, 1);
    lcd.print(pump_work_until);
    check_pump_times();
  }

  // BUTTONS
  button_1_state = digitalRead(button_1);
  button_2_state = digitalRead(button_2);
  button_3_state = digitalRead(button_3);

  if (button_1_state == HIGH && button_1_memstate == LOW) {
    Serial.println("btn_1_down");
    idle_time_count = 0;
    button_1_memstate = button_1_state;
    button_1_holdcount = 0;
    switch (menu_state) {
      case menu_ask_delete_log:
        menu_state = menu_idle;
        break;

      case menu_ask_hours:
        if (set_hours < remain_hours && set_hours < hours_limit) {
          set_hours++;
          break;
        }

      case menu_ask_start:
        menu_state = menu_set_start;
        nfc_write_buffer = get_date_compact();
        remain_hours -= set_hours;
        String br = String(remain_hours);
        for (int i = 0; i < 4 - br.length(); i ++) {
          nfc_write_buffer += "0";
        }
        nfc_write_buffer += br;
        nfc_write_buffer += nfc_card_uid.substring(0, 4);
        Serial.println(nfc_write_buffer);
        break;
    }
  } else {
    if (button_1_state == HIGH) {
      button_1_holdcount++;
      if (button_1_holdcount == 10000) {
        ReadSD();
        lcd.clear();
        lcd.setCursor(0, 0);
        lcd.print("Pump id:");
        lcd.print(pump_id);
        delay(3000);
      }
    }
    button_1_memstate = button_1_state;
  }



  if (button_2_state == HIGH && button_2_memstate == LOW) {
    Serial.println("btn_2_down");
    idle_time_count = 0;
    button_2_memstate = button_2_state;
    button_2_holdcount = 0;
    switch (menu_state) {
      case menu_ask_delete_log:
        menu_state = menu_idle;
        break;
      case menu_ask_hours:
        if (set_hours > 0) {
          set_hours--;
        }
        break;
      case menu_ask_start:
        lcd.clear();
        lcd_shout(F("BgAlTE THN KAPTA"), 0, 0);
        delay(5000);
        menu_state = menu_idle;
    }
  } else {
    if (button_2_state == HIGH) {
      button_2_holdcount++;
      if (button_2_holdcount == 10000) {
        lcd.clear();
        lcd.setCursor(0, 0);
        lcd.print("Delete log?");
        menu_state = menu_ask_delete_log;
      }
    }
    button_2_memstate = button_2_state;
  }

  if (button_3_state == HIGH && button_3_memstate == LOW) {
    Serial.println("btn_3_down");
    idle_time_count = 0;
    button_3_memstate = button_3_state;
    button_3_holdcount = 0;
    if (menu_state == menu_ask_hours) {
      lcd.clear();
      menu_state = menu_ask_start;
    }
  } else {
    if (button_3_state == HIGH) {
      button_3_holdcount++;
      if (button_3_holdcount == 10000) {
        if (menu_state == menu_ask_delete_log) {
          if (!file.remove()) {
            lcd.clear();
            lcd.setCursor(0, 0);
            lcd.print("Delete failed.");
          } else {
            if (!file.open(log_file, O_CREAT | O_WRITE | O_EXCL)) {
              Serial.println(F("  Unable to create file."));
            } else {
              Serial.println(F("  Log file created."));
              log_it(pump_id);   // New file. Write pump id
              delay(1000);
            }
          }
          menu_state = menu_idle;
        }
      }
    }
    button_3_memstate = button_3_state;
  }

  // RTC
  if (RTC.read(tm)) {
    if (menu_state == menu_idle && tm.Second != mem_last_second) {
      get_time_stamp(true);
      mem_last_second = tm.Second;
    }
  } else {
    lcd.clear();
    lcd.setCursor(0, 0);
    if (RTC.chipPresent()) {
      lcd.print(F("RTC error time"));
    } else {
      lcd.print(F("RTC disconnected"));
    }
    delay(5000);
  }



  if (menu_state == menu_show_remain_work) {
    delay(100);
    return;
  }

  bool cp = mfrc522.PICC_IsNewCardPresent();
  delay(10);
  bool nc =  mfrc522.PICC_ReadCardSerial();
  delay(100);

  /*
    Serial.print("menu_state:");
    Serial.print(menu_state);
    Serial.print(" nfc_buffer:");
    Serial.print(nfc_write_buffer.length());
    Serial.print(" PICC_IsNewCardPresent:");
    Serial.print(cp);
    Serial.print(" PICC_ReadCardSerial:");
    Serial.println(nc);
    delay(50);
  */


  // Menu select, no buffer to write, card is the same -> break loop
  // if (nfc_write_buffer.length() == 0 && cp == 0) {
  // delay(loop_delay);
  // return;
  // }

  // Serial.println("card");

  // NFC
  // Look for new cards
  if (nfc_new_init == 255) {
    nfc_write_buffer = "";
    nfc_card_uid = "";
    nfc_new_init = 0;
    fails_count = 0;
    mfrc522.PCD_Reset();
    mfrc522.PCD_Init();
    Serial.println("Reinitialize NFC");
  }

  if (!cp) {
    if (nfc_write_buffer.length() != 16) {
      // Serial.println("NFC No buffer, no card");
      return;
    }
  } else {
    nfc_new_init = 1;
  }

  if (nfc_write_buffer.length() == 16) {
    if (nfc_new_init == 0) {
      mfrc522.PCD_Reset();
      mfrc522.PCD_Init();
      Serial.println("Reinitialize NFC for write");
      fails_count = 0;
      nfc_new_init = 1;
    }
  }

  // Select one of the cards
  if (!nc) {
    fails_count++;
    if (fails_count > 5) {
      Serial.print("Card selection failed");
      lcd.clear();
      lcd_shout("H KAPTA EXEI", 0, 0);
      lcd_shout("AfAIPEuEI...", 0, 1);
      nfc_new_init = 255;
      menu_state = menu_idle;
      delay(5000);
    }
    Serial.println("fail");
    return;
  }


  nfc_new_init = 0;

  Serial.print(F("nfc_uid:"));
  nfc_card_uid = string_byte_array(mfrc522.uid.uidByte, mfrc522.uid.size);
  Serial.print(nfc_card_uid);
  Serial.println("#");


  // Serial.print(F("type:"));
  MFRC522::PICC_Type piccType = mfrc522.PICC_GetType(mfrc522.uid.sak);
  // Serial.print(mfrc522.PICC_GetTypeName(piccType));
  // Serial.println("#");


  Serial.print(F("write:"));
  Serial.print(nfc_write_buffer);
  Serial.println("#");

  // Check for compatibility
  if (piccType != MFRC522::PICC_TYPE_MIFARE_MINI &&  piccType != MFRC522::PICC_TYPE_MIFARE_1K &&  piccType != MFRC522::PICC_TYPE_MIFARE_4K) {
    lcd.clear();
    lcd.setCursor(0, 0);
    lcd.print(F("Unknown NFC type."));
    delay(2000);
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

  bstr = "";
  String data_block = (char*)buffer;
  bstr += data_block;
  bstr += "#";
  Serial.println(bstr);
  String gh = data_block.substring(8, 12);
  remain_hours = gh.toInt();
  Serial.print("Remain hours:");
  Serial.println(remain_hours);
  if (remain_hours == 0) {
    lcd.clear();
    lcd_shout("dEN YpAPXOYN", 0, 0);
    lcd_shout("dIAuEsIMEs vPEs", 0, 1);
    menu_state = menu_idle;
    delay(5000);
    set_hours = 0;
  } else {
    menu_state = menu_ask_hours;
    set_hours = 1;
  }
  lcd.clear();

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
      nfc_new_init = 255;
      Serial.println("card check");
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
      String bstr = "";
      bstr += nfc_card_uid;
      bstr += ";";
      bstr += String(set_hours);
      bstr += ";";
      bstr += get_time_stamp(false);
      log_it(bstr);

      lcd.clear();
      digitalWrite(22, LOW);
      build_pump_work_until();
      menu_state = menu_show_remain_work;
    } else {
      lcd.clear();
      lcd_shout(F("pPOBlHMA..."), 0, 0);
    }
    nfc_card_uid = "";
    nfc_write_buffer = "";
    remain_hours = 0;
    set_hours = 0;
  }

  // Halt PICC
  mfrc522.PICC_HaltA();
  // Stop encryption on PCD
  mfrc522.PCD_StopCrypto1();
  Serial.println("END#");

  delay(100);
}




















// SD
uint8_t cidDmp() {
  cid_t cid;
  if (!sd.card()->readCID(&cid)) {
    sdErrorMsg("readCID failed");
    return false;
  }
  cout << F("\nManufacturer ID: ");
  cout << hex << int(cid.mid) << dec << endl;
  cout << F("OEM ID: ") << cid.oid[0] << cid.oid[1] << endl;
  cout << F("Product: ");
  for (uint8_t i = 0; i < 5; i++) {
    cout << cid.pnm[i];
  }
  cout << F("\nVersion: ");
  cout << int(cid.prv_n) << '.' << int(cid.prv_m) << endl;
  cout << F("Serial number: ") << hex << cid.psn << dec << endl;
  cout << F("Manufacturing date: ");
  cout << int(cid.mdt_month) << '/';
  cout << (2000 + cid.mdt_year_low + 10 * cid.mdt_year_high) << endl;
  cout << endl;
  return true;
}
//------------------------------------------------------------------------------
uint8_t csdDmp() {
  csd_t csd;
  uint8_t eraseSingleBlock;
  if (!sd.card()->readCSD(&csd)) {
    sdErrorMsg("readCSD failed");
    return false;
  }
  if (csd.v1.csd_ver == 0) {
    eraseSingleBlock = csd.v1.erase_blk_en;
    eraseSize = (csd.v1.sector_size_high << 1) | csd.v1.sector_size_low;
  } else if (csd.v2.csd_ver == 1) {
    eraseSingleBlock = csd.v2.erase_blk_en;
    eraseSize = (csd.v2.sector_size_high << 1) | csd.v2.sector_size_low;
  } else {
    cout << F("csd version error\n");
    return false;
  }
  eraseSize++;
  cout << F("cardSize: ") << 0.000512 * cardSize;
  cout << F(" MB (MB = 1,000,000 bytes)\n");

  String bstr;
  lcd.clear();
  lcd.setCursor(0, 0);
  bstr = "Card: ";
  bstr += String(0.000512 * cardSize, 2);
  bstr += "MB";
  lcd.print(bstr);

  cout << F("flashEraseSize: ") << int(eraseSize) << F(" blocks\n");
  cout << F("eraseSingleBlock: ");
  if (eraseSingleBlock) {
    cout << F("true\n");
  } else {
    cout << F("false\n");
  }
  return true;
}
//------------------------------------------------------------------------------
// print partition table
uint8_t partDmp() {
  mbr_t mbr;
  if (!sd.card()->readBlock(0, (uint8_t*)&mbr)) {
    sdErrorMsg("read MBR failed");
    return false;
  }
  for (uint8_t ip = 1; ip < 5; ip++) {
    part_t *pt = &mbr.part[ip - 1];
    if ((pt->boot & 0X7F) != 0 || pt->firstSector > cardSize) {
      cout << F("\nNo MBR. Assuming Super Floppy format.\n");
      return true;
    }
  }
  cout << F("\nSD Partition Table\n");
  cout << F("part,boot,type,start,length\n");
  for (uint8_t ip = 1; ip < 5; ip++) {
    part_t *pt = &mbr.part[ip - 1];
    cout << int(ip) << ',' << hex << int(pt->boot) << ',' << int(pt->type);
    cout << dec << ',' << pt->firstSector << ',' << pt->totalSectors << endl;
  }
  return true;
}
//------------------------------------------------------------------------------
void volDmp() {
  cout << F("\nVolume is FAT") << int(sd.vol()->fatType()) << endl;
  cout << F("blocksPerCluster: ") << int(sd.vol()->blocksPerCluster()) << endl;
  cout << F("clusterCount: ") << sd.vol()->clusterCount() << endl;
  cout << F("freeClusters: ");
  uint32_t volFree = sd.vol()->freeClusterCount();
  cout <<  volFree << endl;
  float fs = 0.000512 * volFree * sd.vol()->blocksPerCluster();
  cout << F("freeSpace: ") << fs << F(" MB (MB = 1,000,000 bytes)\n");

  lcd.setCursor(0, 1);
  String bstr;
  bstr = "Free: ";
  bstr += String(fs, 2);
  bstr += "MB";
  lcd.print (bstr);
  while (digitalRead(button_1)) {
    delay(100);
  }
  bstr = "Log: ";
  bstr += String(file.fileSize());
  lcd.clear();
  lcd.setCursor(0, 0);
  lcd.print (bstr);
  delay(2000);
  while (digitalRead(button_1)) {
    delay(100);
  }


  cout << F("fatStartBlock: ") << sd.vol()->fatStartBlock() << endl;
  cout << F("fatCount: ") << int(sd.vol()->fatCount()) << endl;
  cout << F("blocksPerFat: ") << sd.vol()->blocksPerFat() << endl;
  cout << F("rootDirStart: ") << sd.vol()->rootDirStart() << endl;
  cout << F("dataStartBlock: ") << sd.vol()->dataStartBlock() << endl;
  if (sd.vol()->dataStartBlock() % eraseSize) {
    cout << F("Data area is not aligned on flash!\n");
  }
}
//------------------------------------------------------------------------------

void ReadSD() {
  // SD
  // Wait for USB Serial
  while (!Serial) {
    SysCall::yield();
  }
  uint32_t t = millis();
  // use uppercase in hex and use 0X base prefix
  cout << uppercase << showbase << endl;

  // F stores strings in flash to save RAM
  cout << F("SdFat version: ") << SD_FAT_VERSION << endl;
#if !USE_SDIO
  if (DISABLE_CHIP_SELECT < 0) {
    cout << F(
           "\nAssuming the SD is the only SPI device.\n"
           "Edit DISABLE_CHIP_SELECT to disable another device.\n");
  } else {
    cout << F("\nDisabling SPI device on pin ");
    cout << int(DISABLE_CHIP_SELECT) << endl;
    pinMode(DISABLE_CHIP_SELECT, OUTPUT);
    digitalWrite(DISABLE_CHIP_SELECT, HIGH);
  }
  cout << F("\nAssuming the SD chip select pin is: ") << int(SD_CHIP_SELECT);
  cout << F("\nEdit SD_CHIP_SELECT to change the SD chip select pin.\n");
#endif  // !USE_SDIO

#if USE_SDIO
  if (!sd.cardBegin()) {
    sdErrorMsg("\ncardBegin failed");
    return;
  }
#else  // USE_SDIO
  // Initialize at the highest speed supported by the board that is
  // not over 50 MHz. Try a lower speed if SPI errors occur.
  if (!sd.cardBegin(SD_CHIP_SELECT, SD_SCK_MHZ(50))) {
    sdErrorMsg("cardBegin failed");
    return;
  }
#endif  // USE_SDIO

  t = millis() - t;

  cardSize = sd.card()->cardSize();
  if (cardSize == 0) {
    sdErrorMsg("cardSize failed");
    return;
  }
  cout << F("\ninit time: ") << t << " ms" << endl;
  cout << F("\nCard type: ");
  switch (sd.card()->type()) {
    case SD_CARD_TYPE_SD1:
      cout << F("SD1\n");
      break;

    case SD_CARD_TYPE_SD2:
      cout << F("SD2\n");
      break;

    case SD_CARD_TYPE_SDHC:
      if (cardSize < 70000000) {
        cout << F("SDHC\n");
      } else {
        cout << F("SDXC\n");
      }
      break;

    default:
      cout << F("Unknown\n");
  }
  if (!cidDmp()) {
    return;
  }
  if (!csdDmp()) {
    return;
  }
  uint32_t ocr;
  if (!sd.card()->readOCR(&ocr)) {
    sdErrorMsg("\nreadOCR failed");
    return;
  }
  cout << F("OCR: ") << hex << ocr << dec << endl;
  if (!partDmp()) {
    return;
  }
  if (!sd.fsBegin()) {
    sdErrorMsg("\nFile System initialization failed.\n");
    return;
  }
  volDmp();
}

void log_it(String sstr) {
  Serial.print(F("log_it "));
  Serial.println(sstr);
  file.print(sstr);
  file.print("\r\n");
  file.flush();
  if (!file.sync()) {
    Serial.println(F("File sync error."));
  }
}

String get_time_stamp(bool on_lcd) {
  tmElements_t tm;

  if (RTC.read(tm)) {
    String bstr = "";
    String btime = "";

    bstr += tm.Day < 10 ? "0" : "";
    bstr += tm.Day;
    bstr += "/";
    bstr += tm.Month < 10 ? "0" : "";
    bstr += tm.Month;
    bstr += "/";
    bstr += tmYearToCalendar(tm.Year);

    if (on_lcd == true) {
      lcd.clear();
      lcd.setCursor(0, 0);
      lcd.print(bstr);
    }

    bstr += " ";

    btime  = tm.Hour < 10 ? "0" : "";
    btime += tm.Hour;
    btime += ":";
    btime += tm.Minute < 10 ? "0" : "";
    btime += tm.Minute;
    btime += ":";
    btime += tm.Second < 10 ? "0" : "";
    btime += tm.Second;

    if (on_lcd == true) {
      lcd.setCursor(0, 1);
      lcd.print(btime);
    }

    bstr += btime;

    return bstr;
  } else {
    return F("timedate error.");
  }
}

String get_date_compact() {
  tmElements_t tm;

  if (RTC.read(tm)) {
    String bstr = "";
    bstr += tmYearToCalendar(tm.Year);
    bstr += tm.Month < 10 ? "0" : "";
    bstr += tm.Month;
    bstr += tm.Day < 10 ? "0" : "";
    bstr += tm.Day;
    return bstr;
  } else {
    return F("20010101");
  }
}

void build_pump_work_until() {
  tmElements_t tm;

  RTC.read(tm);

  String bstr = "";
  int newHour, newMinute;

  float csmin = tm.Hour * 60 + tm.Minute;
  float cemin = (tm.Hour + set_hours) * 60 + tm.Minute + 1;

  if (cemin > 1440) {
    newHour = floor((cemin - 1440) / 60);
    newMinute = floor(cemin - floor(cemin / 60) * 60);

    bstr = "AYPIO ";
    bstr += newHour < 10 ? "0" : "";
    bstr += String(newHour);
    bstr += ":";
    bstr += newMinute < 10 ? "0" : "";
    bstr += String(newMinute);
  } else {
    newHour = floor(cemin / 60);
    newMinute = floor(cemin - floor(cemin / 60) * 60);

    bstr = char(0xDA);
    bstr += "HMEPA ";
    bstr += newHour < 10 ? "0" : "";
    bstr += String(newHour);
    bstr += ":";
    bstr += newMinute < 10 ? "0" : "";
    bstr += String(newMinute);
  }

  // DEBUG -> Hours to minutes
  // pump_end_time = tm.Month * 44640 + tm.Day * 1440 + (tm.Hour + set_hours) * 60 + tm.Minute - 59 * set_hours;
  // RELEASE
  pump_end_time = tm.Month * 44640 + tm.Day * 1440 + (tm.Hour + set_hours) * 60 + tm.Minute;
  Serial.println(pump_end_time);
  Serial.println(bstr);
  pump_work_until = bstr;
}

void check_pump_times() {
  if (pump_end_time == 0) {
    return;
  }

  tmElements_t tm;

  if (RTC.read(tm)) {
    long pump_current_time = tm.Month * 44640 + tm.Day * 1440 + tm.Hour * 60 + tm.Minute;
    /*
      Serial.print ("check_pump_times: ");
      Serial.print (pump_current_time);
      Serial.print (", ");
      Serial.println (pump_end_time);
    */
    if (pump_current_time > pump_end_time) {
      Serial.print ("Done");
      digitalWrite(22, HIGH);
      pump_end_time = 0;
      menu_state = menu_idle;
    }
  }
}

// NFC
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

void lcd_shout(String str, byte px, byte py) {
  for (int i = 0; i < str.length(); i++) {
    switch (str[i]) {
      case 'g': str[i] = 0xC9; break;
      case 'd': str[i] = 0x7F; break;
      case 'u': str[i] = 0xD6; break;
      case 'l': str[i] = 0xD7; break;
      case 'j': str[i] = 0xD8; break;
      case 'p': str[i] = 0xD9; break;
      case 's': str[i] = 0xDA; break;
      case 'f': str[i] = 0xDC; break;
      case 'c': str[i] = 0xDD; break;
      case 'v': str[i] = 0xDE; break;
    }
  }
  lcd.setCursor(px, py);
  lcd.print(str);
}




