using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using TapeWorkStation.IO;
using System.Threading;


namespace TapeWorkStation.TapeDevice
{

    #region Typedefenitions
    //this is for compatibility with API types
    using BOOL = System.Int32;
    #endregion
    
    /// <summary>
    /// Low level Tape Operations
    /// </summary>
    public class Tape: IMediaOperator
    {

        #region Types

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TapeGetMediaParameters
        {
            public long Capacity;
            public long Remaining;
            public uint BlockSize;
            public uint PartitionCount;

            public byte IsWriteProtected;
        }

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TapeSetMediaParameters
        {
            public uint BlockSize;
        }

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TapeGetDriveParameters
        {
            public byte ECC;
            public byte driveCompression;
            public byte DataPadding;
            public byte ReportSetMarks;

            public uint DefaultBlockSize;
            public uint MaximumBlockSize;
            public uint MinimumBlockSize;
            public uint PartitionCount;

            public uint FeaturesLow;
            public uint FeaturesHigh;
            public uint EATWarningZone;
        }

        /// <summary>
        /// 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TapeSetDriveParameters
        {
            public byte ECC;
            public byte driveCompression;
            public byte DataPadding;
            public byte ReportSetmarks;

            public uint EOTWarningZoneSize;
        }
        #endregion

        #region Public constants
        /// <summary>
        /// 
        /// </summary>
        public const int FALSE = 0;
        public const int TRUE = 1;
        public const int NULL = 0;

        public const short INVALID_HANDLE_VALUE = -1;

        /// <summary>
        /// file share modes
        /// </summary>
        public const uint FILE_SHARE_READ = 0x00000001;
        public const uint FILE_SHARE_WRITE = 0x00000002;
        public const uint FILE_SHARE_DELETE = 0x00000004;

        /// <summary>
        /// file creation disposition
        /// </summary>
        public const uint CREATE_NEW = 1;
        public const uint CREATE_ALWAYS = 2;
        public const uint OPEN_EXISTING = 3;
        public const uint OPEN_ALWAYS = 4;
        public const uint TRUNCATE_EXISTING = 5;

        /// <summary>
        /// file attributes
        /// </summary>
        public const uint FILE_ATTRIBUTE_READONLY = 0x00000001;
        public const uint FILE_ATTRIBUTE_HIDDEN = 0x00000002;
        public const uint FILE_ATTRIBUTE_SYSTEM = 0x00000004;
        public const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
        public const uint FILE_ATTRIBUTE_ARCHIVE = 0x00000020;
        public const uint FILE_ATTRIBUTE_DEVICE = 0x00000040;
        public const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;
        public const uint FILE_ATTRIBUTE_TEMPORARY = 0x00000100;
        public const uint FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200;
        public const uint FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400;
        public const uint FILE_ATTRIBUTE_COMPRESSED = 0x00000800;
        public const uint FILE_ATTRIBUTE_OFFLINE = 0x00001000;
        public const uint FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000;
        public const uint FILE_ATTRIBUTE_ENCRYPTED = 0x00004000;
        public const uint FILE_ATTRIBUTE_VIRTUAL = 0x00010000;

        /// <summary>
        /// file flags
        /// </summary>
        public const uint FILE_FLAG_WRITE_THROUGH = 0x80000000;
        public const uint FILE_FLAG_OVERLAPPED = 0x40000000;
        public const uint FILE_FLAG_NO_BUFFERING = 0x20000000;
        public const uint FILE_FLAG_RANDOM_ACCESS = 0x10000000;
        public const uint FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000;
        public const uint FILE_FLAG_DELETE_ON_CLOSE = 0x04000000;
        public const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
        public const uint FILE_FLAG_POSIX_SEMANTICS = 0x01000000;
        public const uint FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000;
        public const uint FILE_FLAG_OPEN_NO_RECALL = 0x00100000;
        public const uint FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000;

        /*
        /// <summary>
        /// file security features
        /// </summary>
        public const uint SECURITY_ANONYMOUS
        public const uint SECURITY_CONTEXT_TRACKING
        public const uint SECURITY_DELEGATION
        public const uint SECURITY_EFFECTIVE_ONLY
        public const uint SECURITY_IDENTIFICATION
        public const uint SECURITY_IMPERSONATION
        */

        /// <summary>
        /// file desired access
        /// </summary>
        public const uint GENERIC_READ = 0x80000000;
        public const uint GENERIC_WRITE = 0x40000000;

        /// <summary>
        /// Tape device preparation Operation
        /// </summary>
        public const uint TAPE_LOAD = 0;
        public const uint TAPE_UNLOAD = 1;
        public const uint TAPE_TENSION = 2;
        public const uint TAPE_LOCK = 3;
        public const uint TAPE_UNLOCK = 4;
        public const uint TAPE_FORMAT = 5;

        /// <summary>
        /// return codes
        /// </summary>
        public const uint NO_ERROR = 0;
        public const uint ERROR_BEGINNING_OF_MEDIA = 1102;
        public const uint ERROR_BUS_RESET = 1111;
        public const uint ERROR_DEVICE_NOT_PARTITIONED = 1107;
        public const uint ERROR_END_OF_MEDIA = 1100;
        public const uint ERROR_FILEMARK_DETECTED = 1101;
        public const uint ERROR_INVALID_BLOCK_LENGTH = 1106;
        public const uint ERROR_MEDIA_CHANGED = 1110;
        public const uint ERROR_NO_DATA_DETECTED = 1104;
        public const uint ERROR_NO_MEDIA_IN_DRIVE = 1112;
        public const uint ERROR_NOT_SUPPORTED = 50;
        public const uint ERROR_PARTITION_FAILURE = 1105;
        public const uint ERROR_SETMARK_DETECTED = 1103;
        public const uint ERROR_UNABLE_TO_LOCK_MEDIA = 1108;
        public const uint ERROR_UNABLE_TO_UNLOAD_MEDIA = 1109;
        public const uint ERROR_WRITE_PROTECT = 19;

        /// <summary>
        /// types of positioning
        /// </summary>
        public const uint TAPE_REWIND = 0;
        public const uint TAPE_ABSOLUTE_BLOCK = 1;
        public const uint TAPE_LOGICAL_BLOCK = 2;
        public const uint TAPE_PSEUDO_LOGICAL_BLOCK = 3;
        public const uint TAPE_SPACE_END_OF_DATA = 4;
        public const uint TAPE_SPACE_RELATIVE_BLOCKS = 5;
        public const uint TAPE_SPACE_FILEMARKS = 6;
        public const uint TAPE_SPACE_SEQUENTIAL_FMKS = 7;
        public const uint TAPE_SPACE_SETMARKS = 8;
        public const uint TAPE_SPACE_SEQUENTIAL_SMKS = 9;

        /// <summary>
        /// position type
        /// </summary>
        public const uint TAPE_ABSOLUTE_POSITION = 0;
        public const uint TAPE_LOGICAL_POSITION = 1;
        public const uint TAPE_PSEUDO_LOGICAL_POSITION = 2;

        /// <summary>
        /// get tape parameters operation type
        /// </summary>
        public const uint GET_TAPE_MEDIA_INFORMATION = 0;
        public const uint GET_TAPE_DRIVE_INFORMATION = 1;

        /// <summary>
        /// Drive parameters features low
        /// </summary>
        public const uint TAPE_DRIVE_FIXED = 0x00000001;
        public const uint TAPE_DRIVE_SELECT = 0x00000002;
        public const uint TAPE_DRIVE_INITIATOR = 0x00000004;
        public const uint TAPE_DRIVE_ERASE_SHORT = 0x00000010;
        public const uint TAPE_DRIVE_ERASE_LONG = 0x00000020;
        public const uint TAPE_DRIVE_ERASE_BOP_ONLY = 0x00000040;
        public const uint TAPE_DRIVE_ERASE_IMMEDIATE = 0x00000080;
        public const uint TAPE_DRIVE_TAPE_CAPACITY = 0x00000100;
        public const uint TAPE_DRIVE_TAPE_REMAINING = 0x00000200;
        public const uint TAPE_DRIVE_FIXED_BLOCK = 0x00000400;
        public const uint TAPE_DRIVE_VARIABLE_BLOCK = 0x00000800;
        public const uint TAPE_DRIVE_WRITE_PROTECT = 0x00001000;
        public const uint TAPE_DRIVE_EOT_WZ_SIZE = 0x00002000;
        public const uint TAPE_DRIVE_ECC = 0x00010000;
        public const uint TAPE_DRIVE_COMPRESSION = 0x00020000;
        public const uint TAPE_DRIVE_PADDING = 0x00040000;
        public const uint TAPE_DRIVE_REPORT_SMKS = 0x00080000;
        public const uint TAPE_DRIVE_GET_ABSOLUTE_BLK = 0x00100000;
        public const uint TAPE_DRIVE_GET_LOGICAL_BLK = 0x00200000;
        public const uint TAPE_DRIVE_SET_EOT_WZ_SIZE = 0x00400000;
        public const uint TAPE_DRIVE_EJECT_MEDIA = 0x01000000;
        public const uint TAPE_DRIVE_CLEAN_REQUESTS = 0x02000000;
        public const uint TAPE_DRIVE_SET_CMP_BOP_ONLY = 0x04000000;

        /// <summary>
        /// Drive parameters features high
        /// </summary>
        public const uint TAPE_DRIVE_LOAD_UNLOAD = 0x80000001;
        public const uint TAPE_DRIVE_TENSION = 0x80000002;
        public const uint TAPE_DRIVE_LOCK_UNLOCK = 0x80000004;
        public const uint TAPE_DRIVE_REWIND_IMMEDIATE = 0x80000008;
        public const uint TAPE_DRIVE_SET_BLOCK_SIZE = 0x80000010;
        public const uint TAPE_DRIVE_LOAD_UNLD_IMMED = 0x80000020;
        public const uint TAPE_DRIVE_TENSION_IMMED = 0x80000040;
        public const uint TAPE_DRIVE_LOCK_UNLK_IMMED = 0x80000080;
        public const uint TAPE_DRIVE_SET_ECC = 0x80000100;
        public const uint TAPE_DRIVE_SET_COMPRESSION = 0x80000200;
        public const uint TAPE_DRIVE_SET_PADDING = 0x80000400;
        public const uint TAPE_DRIVE_SET_REPORT_SMKS = 0x80000800;
        public const uint TAPE_DRIVE_ABSOLUTE_BLK = 0x80001000;
        public const uint TAPE_DRIVE_ABS_BLK_IMMED = 0x80002000;
        public const uint TAPE_DRIVE_LOGICAL_BLK = 0x80004000;
        public const uint TAPE_DRIVE_LOG_BLK_IMMED = 0x80008000;
        public const uint TAPE_DRIVE_END_OF_DATA = 0x80010000;
        public const uint TAPE_DRIVE_RELATIVE_BLKS = 0x80020000;
        public const uint TAPE_DRIVE_FILEMARKS = 0x80040000;
        public const uint TAPE_DRIVE_SEQUENTIAL_FMKS = 0x80080000;
        public const uint TAPE_DRIVE_SETMARKS = 0x80100000;
        public const uint TAPE_DRIVE_SEQUENTIAL_SMKS = 0x80200000;
        public const uint TAPE_DRIVE_REVERSE_POSITION = 0x80400000;
        public const uint TAPE_DRIVE_SPACE_IMMEDIATE = 0x80800000;
        public const uint TAPE_DRIVE_WRITE_SETMARKS = 0x81000000;
        public const uint TAPE_DRIVE_WRITE_FILEMARKS = 0x82000000;
        public const uint TAPE_DRIVE_WRITE_SHORT_FMKS = 0x84000000;
        public const uint TAPE_DRIVE_WRITE_LONG_FMKS = 0x88000000;
        public const uint TAPE_DRIVE_WRITE_MARK_IMMED = 0x90000000;
        public const uint TAPE_DRIVE_FORMAT = 0xA0000000;
        public const uint TAPE_DRIVE_FORMAT_IMMEDIATE = 0xC0000000;

        /// <summary>
        /// set tape parameters operation type
        /// </summary>
        public const uint SET_TAPE_MEDIA_INFORMATION = 0;
        public const uint SET_TAPE_DRIVE_INFORMATION = 1;

        /// <summary>
        /// partition method
        /// </summary>
        public const uint TAPE_FIXED_PARTITIONS = 0;
        public const uint TAPE_SELECT_PARTITIONS = 1;
        public const uint TAPE_INITIATOR_PARTITIONS = 2;

        /// <summary>
        /// Erase type
        /// </summary>
        public const uint TAPE_ERASE_SHORT = 0;
        public const uint TAPE_ERASE_LONG = 1;

        /// <summary>
        /// tapemark types
        /// </summary>
        public const uint TAPE_SETMARKS = 0;
        public const uint TAPE_FILEMARKS = 1;
        public const uint TAPE_SHORT_FILEMARKS = 2;
        public const uint TAPE_LONG_FILEMARKS = 3;
        #endregion

        #region PInvoke

        /// <summary>
        /// CreateFile Win API
        /// </summary>
        /// <param name="lpFileName"></param>
        /// <param name="dwDesiredAccess"></param>
        /// <param name="dwShareMode"></param>
        /// <param name="lpSecurityAttributes"></param>
        /// <param name="dwCreationDisposition"></param>
        /// <param name="dwFlagsAndAttributes"></param>
        /// <param name="hTemplateFile"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile
            );

        /// <summary>
        /// PrepareTape Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="prepareType"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint PrepareTape(
            SafeFileHandle handle,
            uint prepareType,
            BOOL isImmediate
            );

        /// <summary>
        /// SetTapePosition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="positionMethod"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint SetTapePosition(
            SafeFileHandle handle,
            uint positionMethod,
            uint partition,
            uint offsetLow,
            uint offsetHigh,
            BOOL isImmediate
            );

        /// <summary>
        /// GetTapePosition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="positionType"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapePosition(
            SafeFileHandle handle,
            uint positionType,
            ref uint partition,
            ref uint offsetLow,
            ref uint offsetHigh
            );

        /// <summary>
        /// GetTapeParameters Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="operationType"></param>
        /// <param name="size"></param>
        /// <param name="mediaOrDriveInfo"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapeParameters(
           SafeFileHandle handle,
           uint operationType,
           ref uint size,
           IntPtr mediaOrDriveInfo
           );

        /// <summary>
        /// SetTapeParameters Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="operationType"></param>
        /// <param name="mediaOrDriveInfo"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint SetTapeParameters(
           SafeFileHandle handle,
           uint operationType,
           IntPtr mediaOrDriveInfo
            );

        /// <summary>
        /// CreateTapePartition Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="PartitionMethod"></param>
        /// <param name="count"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint CreateTapePartition(
            SafeFileHandle handle,
            uint PartitionMethod,
            uint count,
            uint size
            );

        /// <summary>
        /// EraseTape Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="eraseType"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint EraseTape(
        SafeFileHandle handle,
        uint eraseType,
        BOOL isImmediate
        );

        /// <summary>
        /// WriteTapemark Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="tapemarkType"></param>
        /// <param name="tapemarkCount"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint WriteTapemark(
        SafeFileHandle handle,
        uint tapemarkType,
        uint tapemarkCount,
        BOOL isImmediate
        );

        /// <summary>
        /// GetTapeStatus Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetTapeStatus(
        SafeFileHandle handle
        );

        /// <summary>
        /// ReadFile Win API
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="buffer"></param>
        /// <param name="numberOfBytesToRead"></param>
        /// <param name="numberOfBytesRead"></param>
        /// <param name="overlappedBuffer"></param>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern bool ReadFile(
        SafeFileHandle handle,
        IntPtr buffer,
        uint numberOfBytesToRead,
        ref uint numberOfBytesRead,
        IntPtr overlappedBuffer
        );

        /// <summary>
        /// GetLastError Win API
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern uint GetLastError();

        #endregion

        #region Private variables

        private SafeFileHandle m_handleDeviceValue = null;

        private Nullable<TapeGetDriveParameters> m_driveInfo = null;
        private Nullable<TapeGetMediaParameters> m_mediaInfo = null;

        private uint m_tapeDriveNumber;
        private string m_tapeDriveName;

        #endregion

        #region Public methods
        /// <summary>
        /// Attempts to open tape drive
        /// </summary>
        /// <param name="strTapeDriveName"></param>
        public bool Open(string strTapeDriveName)
        {
            // initialize private member with passed string 
            m_tapeDriveName = strTapeDriveName;

            // try to open device
            m_handleDeviceValue = CreateFile(
                            m_tapeDriveName,
                            GENERIC_READ | GENERIC_WRITE,
                            0,
                            IntPtr.Zero,
                            OPEN_EXISTING,
                            FILE_ATTRIBUTE_NORMAL,
                            IntPtr.Zero
                            );

            if (m_handleDeviceValue.IsInvalid)
            {
                // could not open
                m_tapeDriveName = null;

                return false;
            }

            // initialize private member
            m_tapeDriveNumber = uint.Parse(strTapeDriveName.Remove(0, 8));

            this.GetTapeStatus();// this is to reset ERROR_MEDIA_CHANGED status
            Thread.Sleep(500);

            return true;
        }

        /// <summary>
        /// Attempts to open tape drive
        /// </summary>
        /// <param name="nTapeDevice"></param>
        public bool Open(uint nTapeDevice)
        {
            return Open("\\\\.\\TAPE" + nTapeDevice);
        }

        /// <summary>
        /// Loads tape
        /// </summary>
        public bool Load(ref uint returnCode)
        {
            return Prepare(TAPE_LOAD, ref returnCode, TRUE);
        }

        /// <summary>
        /// Unloads tape
        /// </summary>
        public bool Unload(ref uint returnCode)
        {
            return Prepare(TAPE_UNLOAD, ref returnCode, TRUE);
        }

        /// <summary>
        /// Locks tape in a drive
        /// </summary>
        public bool Lock(ref uint returnCode)
        {
            return Prepare(TAPE_LOCK, ref returnCode, TRUE);
        }

        /// <summary>
        /// Unlocks tape in a drive
        /// </summary>
        public bool Unlock(ref uint returnCode)
        {
            return Prepare(TAPE_UNLOCK, ref returnCode, TRUE);
        }

        /// <summary>
        /// Formats tape
        /// </summary>
        public bool FormatTape(ref uint returnCode)
        {
            return Prepare(TAPE_FORMAT, ref returnCode, TRUE);
        }

        /// <summary>
        /// Adjusts tape tension
        /// </summary>
        public bool AdjustTension(ref uint returnCode)
        {
            return Prepare(TAPE_TENSION, ref returnCode, TRUE);
        }

        /// <summary>
        /// Rewinds tape to BOD
        /// </summary>
        public bool Rewind(ref uint returnCode)
        {
            return SetTapePosition(TAPE_REWIND, 0, 0, 0, FALSE, ref returnCode);
        }

        /// <summary>
        /// Rewinds tape to EOD
        /// </summary>
        public bool SeekToEOD(ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_END_OF_DATA, 0, 0, 0, FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning tape to absolute block
        /// </summary>
        public bool SeekToAbsoluteBlock(long blockAddress, ref uint returnCode)
        {
            return SetTapePosition(TAPE_ABSOLUTE_BLOCK, 0, (uint)blockAddress, (uint)(blockAddress >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning tape to logical block
        /// </summary>
        public bool SeekToLogicalBlock(long blockAddress, ref uint returnCode)
        {
            return SetTapePosition(TAPE_LOGICAL_BLOCK, 0, (uint)blockAddress, (uint)(blockAddress >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning (backward or forward) by spacing number of filemarks
        /// </summary>
        public bool SpaceFileMarks(long filemarksToSpace, ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_FILEMARKS, 0, (uint)filemarksToSpace, (uint)(filemarksToSpace >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Positioning (backward or forward) by spacing number of blocks
        /// </summary>
        public bool SpaceBlocks(long blocksToSpace, ref uint returnCode)
        {
            return SetTapePosition(TAPE_SPACE_RELATIVE_BLOCKS, 0, (uint)blocksToSpace, (uint)(blocksToSpace >> 32), FALSE, ref returnCode);
        }

        /// <summary>
        /// Reads from the tape
        /// </summary>
        public bool Read(ref byte[] buffer, uint bytesToRead, ref uint bytesRead, ref uint returnCode)
        {
            returnCode = NO_ERROR;

            GCHandle h = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            IntPtr ptrBuffer = Marshal.UnsafeAddrOfPinnedArrayElement(buffer, 0);
            h.Free();

            if (ReadFile(m_handleDeviceValue, ptrBuffer, bytesToRead, ref bytesRead, IntPtr.Zero))
            {
                return true;
            }
            //returnCode = GetLastError();
            returnCode = (uint)Marshal.GetLastWin32Error();

            return false;
        }

        /// <summary>
        /// Closes device handle
        /// </summary>
        public void Close()
        {
            if (m_handleDeviceValue != null &&
                !m_handleDeviceValue.IsInvalid &&
                !m_handleDeviceValue.IsClosed)
            {
                m_handleDeviceValue.Close();
                m_tapeDriveName = null;
                m_tapeDriveNumber = 0;
            }
        }

        /// <summary>
        /// Erase tape
        /// </summary>
        public bool Erase(uint eraseType, ref uint returnCode)
        {
            returnCode = EraseTape(m_handleDeviceValue, eraseType, FALSE);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// Writes the number of tape marks
        /// </summary>
        public bool WriteTapemark(uint marksCount, uint marksType, ref uint returnCode)
        {
            returnCode = WriteTapemark(m_handleDeviceValue, marksType, marksCount, FALSE);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// returns tape drive status
        /// </summary>
        public uint GetTapeStatus()
        {
            return GetTapeStatus(m_handleDeviceValue);
        }

        /// <summary>
        /// Returns current tape's position in logical or absolute blocks
        /// </summary>
        public bool GetTapePosition(uint positionType, ref uint partitionNumber, ref uint offsetLow, ref uint offsetHigh, ref uint returnCode)
        {
            returnCode = GetTapePosition(m_handleDeviceValue, positionType, ref partitionNumber, ref offsetLow, ref offsetHigh);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        public bool GetTapePosition(ref long Position, ref uint returnCode)
        {
            uint partition=0;
            uint offsetLow=0;
            uint offsetHigh=0;

            returnCode = GetTapePosition(
                m_handleDeviceValue,
                TAPE_LOGICAL_POSITION,
                ref partition,
                ref offsetLow,
                ref offsetHigh);
            Position = (long)(offsetHigh * Math.Pow(2, 32) + offsetLow);
            if (returnCode == NO_ERROR) return true;

            return false;
        }

        public bool SetBlockSize(uint blockSize, ref uint returnCode)
        {
            IntPtr ptr = IntPtr.Zero;

            TapeSetMediaParameters mediaInfo = new TapeSetMediaParameters();

            mediaInfo.BlockSize = blockSize;

            // Allocate unmanaged memory
            uint size = (uint)Marshal.SizeOf(mediaInfo);
            ptr = Marshal.AllocHGlobal((int)size);

            Marshal.StructureToPtr(mediaInfo, ptr, false);

            returnCode = SetTapeParameters(m_handleDeviceValue, SET_TAPE_MEDIA_INFORMATION, ptr);

            Marshal.FreeHGlobal(ptr);

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns Device number
        /// </summary>
        public uint TapeDriveNumber
        {
            get { return m_tapeDriveNumber; }
        }

        /// <summary>
        /// Returns string with the full device name/path
        /// </summary>
        public string TapeDriveName
        {
            get { return m_tapeDriveName; }
        }

        /// <summary>
        /// Returns Device Handle
        /// </summary>
        public SafeFileHandle HandleDeviceValue
        {
            get { return m_handleDeviceValue; }
        }


        /// <summary>
        /// Returns default block size for current
        /// device
        /// </summary>
        // This has been implemented by original Author
        // Alex has nothing to do with this
        public uint BlockSize
        {
            get
            {
                IntPtr ptr = IntPtr.Zero;
                try
                {
                    if (!m_driveInfo.HasValue)
                    {
                        m_driveInfo = new TapeGetDriveParameters();

                        // Allocate unmanaged memory
                        uint size = (uint)Marshal.SizeOf(m_driveInfo);
                        ptr = Marshal.AllocHGlobal((int)size);

                        Marshal.StructureToPtr(
                            m_driveInfo,
                            ptr,
                            false
                        );




                        uint result = 0;
                        if ((result = GetTapeParameters(
                            m_handleDeviceValue,
                            GET_TAPE_DRIVE_INFORMATION,
                            ref size,
                            ptr)) != NO_ERROR)
                        {
                            throw new TapeWin32Exception(
                                "GetTapeParameters", Marshal.GetLastWin32Error());
                        }

                        // Get managed media Info
                        m_driveInfo = (TapeGetDriveParameters)
                            Marshal.PtrToStructure(ptr, typeof(TapeGetDriveParameters));
                    }


                    return m_driveInfo.Value.DefaultBlockSize;
                }
                finally
                {
                    if (ptr != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(ptr);

                    }
                }
            }
        }
        #endregion

        #region Private methods
        
        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="prepareType"></param>
        /// <param name="returnCode"></param>
        /// <param name="isImmediate"></param>
        /// <returns></returns>
        private bool Prepare(uint prepareType, ref uint returnCode, BOOL isImmediate)
        {
            returnCode = PrepareTape(
                m_handleDeviceValue,
                prepareType,
                isImmediate
                );

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="positionMethod"></param>
        /// <param name="partition"></param>
        /// <param name="offsetLow"></param>
        /// <param name="offsetHigh"></param>
        /// <param name="isImmediate"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool SetTapePosition(uint positionMethod, uint partition, uint offsetLow, uint offsetHigh, BOOL isImmediate, ref uint returnCode)
        {
            returnCode = SetTapePosition(
                m_handleDeviceValue,
                positionMethod,
                partition,
                offsetLow,
                offsetHigh,
                isImmediate
                );

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool SetDriveParameters(TapeSetDriveParameters parameters, ref uint returnCode)
        {
            GCHandle h = GCHandle.Alloc(parameters, GCHandleType.Pinned);
            IntPtr ptrParameters = h.AddrOfPinnedObject();
            h.Free();

            returnCode = SetTapeParameters(
            m_handleDeviceValue,
            SET_TAPE_DRIVE_INFORMATION,
            ptrParameters
            );

            if (returnCode == NO_ERROR) return true;

            return false;
        }

        /// <summary>
        /// This method envelopes API call
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="returnCode"></param>
        /// <returns></returns>
        private bool SetMediaParameters(TapeSetMediaParameters parameters, ref uint returnCode)
        {
            GCHandle h = GCHandle.Alloc(parameters, GCHandleType.Pinned);
            IntPtr ptrParameters = h.AddrOfPinnedObject();
            h.Free();

            returnCode = SetTapeParameters(
            m_handleDeviceValue,
            SET_TAPE_DRIVE_INFORMATION,
            ptrParameters
            );

            if (returnCode == NO_ERROR) return true;

            return false;
        }
        #endregion

        public string ConvertErrCode(uint errcode)
        {
            string errorMessage = errcode.ToString() + ": " + new Win32Exception((int)errcode).Message;
        
            return errorMessage;
        }

    }

    /// <summary>
    /// Exception that will be thrown by tape
    /// operator when one of WIN32 APIs terminates 
    /// with error code 
    /// </summary>
    // This has been implemented by original Author
    // Alex has nothing to do with this
    public class TapeWin32Exception : ApplicationException
    {
        public TapeWin32Exception(string methodName, int win32ErroCode) :
            base(string.Format(
               "WIN32 API method failed : {0} failed with error code {1}",
               methodName,
               win32ErroCode
           )) { }
    }
    /*
        public class TapeStream: MediaStream
        {
            private Tape _dev;


            public TapeStream (string dev)
            {
                _dev = new Tape();
                _dev.Open(dev);
            }

     

            public override IMediaOperator GetMediaDevice()
            {
                return _dev;
            }
            public override int GetMediaDeviceNumber()
            {
                return 0;
            }

            public override int Read(ref byte[] buffer, int length, ref int ret)
            {
                _dev.Read();
            }
            public override int Write(ref byte[] buffer, int length, ref int ret)
            {

            }
            public override int Seek(int seekMethod, int seekArg, Int64 seekOff)
            {

            }
    
        }
     */
}