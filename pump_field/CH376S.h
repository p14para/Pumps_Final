// SYNCHRONIZE
#define CH376S_SER_SYNC_CODE1           0x57
#define CH376S_SER_SYNC_CODE2           0xAB

// COMMANDS
#define CH376S_CMD_GET_IC_VER           0x01    // Obtain chip and firmware version number
#define CH376S_CMD_SET_BAUDRATE         0x02    // Set serial communication baud rate
#define CH376S_CMD_ENTER_SLEEP          0x03    // Go to low-power and suspending
#define CH376S_CMD_RESET_ALL            0x05    // Execute hardware reset
#define CH376S_CMD_CHECK_EXIST          0x06    // Test communication interface and working status
#define CH376S_CMD_SET_SD0_INT          0x0B    // Set interrupt mode of SD0 in SPI
#define CH376S_CMD_GET_FILE_SIZE        0x0C    // Get the current file length
#define CH376S_CMD_SET_USB_MODE         0x15    // Configure the work mode of USB
#define CH376S_CMD_GET_STATUS           0x22    // Get interruption status and cancel requirement
#define CH376S_CMD_RD_USB_DATA0         0x27    // Read data from current interrupt port buffer of USB or receive buffer of host port
#define CH376S_CMD_WR_USB_DATA          0x2C    // Write data to transfer buffer of USB host
#define CH376S_CMD_WR_REQ_DATA          0x2D    // Write requested data block to internal appointed buffer
#define CH376S_CMD_WR_OFS_DATA          0x2E    // Write data block to internal buffer with appointed excursion address
#define CH376S_CMD_SET_FILE_NAME        0x2F    // Set the file name which will be operated
#define CH376S_CMD_DISK_CONNECT         0x30    // (irq) Check the disk connection status
#define CH376S_CMD_DISK_MOUNT           0x31    // (irq) Initialize disk and test disk ready
#define CH376S_CMD_FILE_OPEN            0x32    // (irq) Open file or catalog, enumerate file and catalog
#define CH376S_CMD_FILE_ENUM_GO         0x33    // (irq) Go on to enumerate file and catalog
#define CH376S_CMD_FILE_CREATE          0x34    // (irq) Create file
#define CH376S_CMD_FILE_ERASE           0x35    // (irq) Delete file
#define CH376S_CMD_FILE_CLOSE           0x36    // (irq) Close the open file or catalog
#define CH376S_CMD_DIR_INF0_READ        0x37    // (irq) Read the catalog information of file
#define CH376S_CMD_DIR_INF0_SAVE        0x38    // (irq) Save catalog information of file
#define CH376S_CMD_BYTE_LOCATE          0x39    // (irq) Move the file pointer take byte as unit
#define CH376S_CMD_BYTE_READ            0x3A    // (irq) Read data block from current location take byte as unit
#define CH376S_CMD_BYTE_RD_GO           0x3B    // (irq) Continue byte read
#define CH376S_CMD_BYTE_WRITE           0x3C    // (irq) Write data block from current location take unit as unit
#define CH376S_CMD_BYTE_WR_GO           0x3D    // (irq) Continue byte write
#define CH376S_CMD_DISK_CAPACITY        0x3E    // (irq) Check disk physical capacity
#define CH376S_CMD_DISK_QUERY           0x3F    // (irq) Check disk space
#define CH376S_CMD_DIR_CREATE           0x40    // (irq) Create catalog and open or open the existed catalog
#define CH376S_CMD_SEG_LOCATE           0x4A    // (irq) Move file pointer from current location take fan as unit
#define CH376S_CMD_SEC_READ             0x4B    // (irq) Read data block from current location take fan as unit
#define CH376S_CMD_SEC_WRITE            0x4C    // (irq) Write data block from current location take fan as unit
#define CH376S_CMD_DISK_BOC_CMD         0x50    // (irq) Execute B0 transfer protocol command to USB Storage
#define CH376S_CMD_DISK_READ            0x54    // (irq) Read physical fan from USB storage device
#define CH376S_CMD_DISK_RD_GO           0x55    // (irq) Go on reading operation of USB storage device
#define CH376S_CMD_DISK_WRITE           0x56    // (irq) Write physical fan to USB storage device
#define CH376S_CMD_DISK_WR_GO           0x57    // (irq) Go on writing operation of USB storage device

// USB STATUS (CH376S_CMD_GET_STATUS)
#define CH376S_ST_USB_INT_SUCCESS       0x14    // Success of USB transaction or transfer operation
#define CH376S_ST_USB_INT_CONNECT       0x15    // Detection of USB device attachment
#define CH376S_ST_USB_INT_DISCONNECT    0x16    // Detection of USB device detachment
#define CH376S_ST_USB_INT_BUF_OVER      0x17    // Data error or Buffer overflow
#define CH376S_ST_USB_INT_USB_READY     0x18    // USB device has initialized (appointed USB address)
#define CH376S_ST_USB_INT_DISK_READ     0x1D    // Read operation of storage device
#define CH376S_ST_USB_INT_DISK_WRITE    0x1E    // Write operation of storage device
#define CH376S_ST_USB_INT_DISK_ERR      0x1F    // Failure of USB storage device

// ERROR STATES
#define CH376S_ERR_OPEN_DIR             0x41    // Open directory address is appointed
#define CH376S_ERR_MISS_FILE            0x42    // File doesn’t be found which address is appointed, maybe the name is error
#define CH376S_ERR_FOUND_NAME           0x43    // Search suited file name, or open directory according the request, open file in actual
#define CH376S_ERR_DISK_DISCON          0x82    // Disk doesn’t connect, maybe the disk has cut down
#define CH376S_ERR_LARGE_SECTOR         0x84    // Fan is too big, only support 512 bytes
#define CH376S_ERR_TYPE_ERROR           0x92    // Disk partition doesn’t support, re-prartition by tool
#define CH376S_ERR_BPB_ERROR            0xA1    // Disk doesn’t format, or parameter is errot, re-formate by WINDOWS with default parameter
#define CH376S_ERR_DISK_FULL            0xB1    // File in disk is full, spare space is too small
#define CH376S_ERR_FDT_OVER             0xB2    // Many file in directory, no spare directory, clean up the disk, less than 512 in FAT12/FAT16 root directory
#define CH376S_ERR_FILE_CLOSE           0xB4    // File is closed, re-open file if need

// FAT FILE ATTRIBUTES
#define CH376S_FILE_ATTR_READ_ONLY      0x01
#define CH376S_FILE_ATTR_HIDDEN         0x02
#define CH376S_FILE_ATTR_SYSTEM         0x04
#define CH376S_FILE_ATTR_VOLUME_ID      0x08
#define CH376S_FILE_ATTR_DIRECTORY      0x10
#define CH376S_FILE_ATTR_ARCHIVE        0x20
#define CH376S_FILE_ATTR_LONG_NAME      (CH376S_FILE_ATTR_READ_ONLY | CH376S_FILE_ATTR_HIDDEN | CH376S_FILE_ATTR_SYSTEM | CH376S_FILE_ATTR_VOLUME_ID)
#define CH376S_FILE_ATTR_LONG_NAME_MASK (CH376S_FILE_ATTR_LONG_NAME | CH376S_FILE_ATTR_DIRECTORY | CH376S_FILE_ATTR_ARCHIVE)

// USB MODE SUB COMMANDS
#define CH376S_DEF_USB_PID_SOF          0x05
#define CH376S_DEF_USB_PID_SOF_AUTO     0x06
#define CH376S_DEF_USB_BUS_RESET        0x07

// VARIABLES
#define CH376S_VAR_FILE_SIZE            0x68

#define CH376S_FILE_CLOSE_LEAVE_SIZE    0x00
#define CH376S_FILE_CLOSE_UPDATE_SIZE   0x01

// RETURN STATES
#define CH376S_CMD_RET_SUCCESS          0x51
#define CH376S_CMD_RET_ABORT            0x5F