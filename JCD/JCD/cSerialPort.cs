////////////////////////////////////////////////////////////////////////////////////////////////////
// file:	cSerialPort.cs
//
// summary:	Implements the serial port class
////////////////////////////////////////////////////////////////////////////////////////////////////

/// <summary>   The system. </summary>
using System;
/// <summary>   The system. io. ports. </summary>
using System.IO.Ports;
/// <summary>   The system. windows. forms. </summary>
using System.Windows.Forms;

/// <summary>   A configuration serial port. </summary>
public class ConfigSerialPort
{
    /// <summary>   Name of the port. </summary>
    public string sPortName = "COM3";
    /// <summary>   Zero-based index of the baud rate. </summary>
    public int iBaudRate = 9600;
    /// <summary>   Zero-based index of the data bits. </summary>
    public int iDataBits = 8;
    /// <summary>   The parity. </summary>
    public Parity eParity = Parity.Odd;
    /// <summary>   The stop bits. </summary>
    public StopBits eStopBits = StopBits.One;
    /// <summary>   Size of the read buffer. </summary>
    public int iReadBufferSize = 1024;
    /// <summary>   Size of the write buffer. </summary>
    public int iWriteBufferSize = 1024;
    /// <summary>   true if dtr enable. </summary>
    public bool bDtrEnable = false;
    /// <summary>   true if RTS enable. </summary>
    public bool bRtsEnable = false;
    /// <summary>   Zero-based index of the read time out. </summary>
    public int iReadTimeOut = 1000;
    /// <summary>   Zero-based index of the write time out. </summary>
    public int iWriteTimeOut = 1000;
}

/// <summary>   A serial port. </summary>
public class cSerialPort
{
    /// <summary>   The serial. </summary>
    private SerialPort m_Serial = new SerialPort();
    /// <summary>   The configuration serial. </summary>
    private ConfigSerialPort m_ConfigSerial = new ConfigSerialPort();

    private System.IO.StreamWriter m_LogFile = null;

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void Init(ConfigSerialPort oConfig)</signature>
    ///
    /// <summary> Initialises this object.</summary>
    ///
    /// <param name="oConfig" type="ConfigSerialPort"> The configuration. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void Init(ConfigSerialPort oConfig)
    {
        // Recopie de l'objet de configuration de la liaison série
        m_ConfigSerial = oConfig;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void Open(out Boolean bIsSerialPortAlreadyOpen, out Boolean bErrorOccured,
    ///             out Int32 iErrorCode, out String sErrorMessage)</signature>
    ///
    /// <summary> Opens.</summary>
    ///
    /// <param name="bIsSerialPortAlreadyOpen" type="out Boolean"> [in,out] The error occured. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void Open(out Boolean bIsSerialPortAlreadyOpen, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;

        // Initialisation du booléen indiquant l'état du port série
        bIsSerialPortAlreadyOpen = false;
        // Lancement du bloc try catch
        try
        {
            // Récupération de l'état du port série
            if (m_Serial.IsOpen)
            {
                bIsSerialPortAlreadyOpen = true;
                return;
            }

            // Recopie de la désignation du port
            m_Serial.PortName = m_ConfigSerial.sPortName;
            // Recopie de la vitesse de transfert
            m_Serial.BaudRate = m_ConfigSerial.iBaudRate;
            // Recopie du nombre de bits de data
            m_Serial.DataBits = m_ConfigSerial.iDataBits;
            // Recopie du nombre de bits de parité
            m_Serial.Parity = m_ConfigSerial.eParity;
            // Recopie du nombre de bits de stop
            m_Serial.StopBits = m_ConfigSerial.eStopBits;
            // Recopie de la taille du buffer de lecture
            m_Serial.ReadBufferSize = m_ConfigSerial.iReadBufferSize;
            // Recopie de la taille du buffer d'écriture
            m_Serial.WriteBufferSize = m_ConfigSerial.iWriteBufferSize;
            m_Serial.DtrEnable = m_ConfigSerial.bDtrEnable;
            m_Serial.RtsEnable = m_ConfigSerial.bRtsEnable;
            m_Serial.ReadTimeout = m_ConfigSerial.iReadTimeOut;
            m_Serial.WriteTimeout = m_ConfigSerial.iWriteTimeOut;

            // Ouverture du port série
            m_Serial.Open();

#if DEBUG
            //m_LogFile = new System.IO.StreamWriter(@"C:\SDS\" + m_ConfigSerial.sPortName + "_IBETH.txt");
            //m_LogFile.WriteLine("============ START SESSION ============");
#endif
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void Close(out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)</signature>
    ///
    /// <summary> Closes.</summary>
    ///
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void Close(out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;

        try
        {
            if (m_Serial.IsOpen)
            {
#if DEBUG
                m_LogFile.WriteLine("============ END SESSION ============");
                m_LogFile.Close();
#endif
                m_Serial.Close();
            }              
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void ClearSerial(out Boolean bErrorOccured, out Int32 iErrorCode,
    ///             out String sErrorMessage)</signature>
    ///
    /// <summary> Clears the serial.</summary>
    ///
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void ClearSerial(out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;

        try
        {
            m_Serial.DiscardInBuffer();
            m_Serial.DiscardOutBuffer();
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> private void WriteSerial(String sCommande, out Boolean bErrorOccured,
    ///             out Int32 iErrorCode, out String sErrorMessage)</signature>
    ///
    /// <summary> Writes a serial.</summary>
    ///
    /// <param name="sCommande" type="String">          The commande. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    private void WriteSerial(String sCommande, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;

        try
        {
            m_Serial.DiscardInBuffer();
            m_Serial.DiscardOutBuffer();

            m_Serial.Write(sCommande);
#if DEBUG
            m_LogFile.WriteLine("WriteSerial(" + sCommande.Replace("\r", "\\r").Replace("\n", "\\n") + ")");
#endif
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> private void ReadSerial(Double dTimeout, out String sBufferRead,
    ///             out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)</signature>
    ///
    /// <summary> Reads a serial.</summary>
    ///
    /// <param name="dTimeout" type="Double">           The timeout. </param>
    /// <param name="sBufferRead" type="out String">    [in,out] The buffer read. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ///
    /// <exception cref="TimeoutException"> Thrown when a Timeout error condition occurs. </exception>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void ReadSerial(Double dTimeout, out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;
        sBufferRead = String.Empty;

        try
        {
            int iSizeBuffer = m_ConfigSerial.iReadBufferSize;
            char[] arr = new char[iSizeBuffer];
            sBufferRead = "";
            bool bStop = false;
            TimeSpan ts = TimeSpan.MinValue;
            DateTime dtStart = DateTime.Now;
            int iNbData = 0;
            do
            {
                try
                {
                    iNbData = m_Serial.Read(arr, 0, iSizeBuffer - 1);
                    for (int i = 0; i < iNbData; i++)
                    {
                        sBufferRead += arr[i];
                    }
                }
                catch (TimeoutException)
                {
                    iNbData = 0;
                }

                //Gestion timeout
                ts = DateTime.Now - dtStart;
                if (ts.TotalMilliseconds > dTimeout)
                {
                    throw new TimeoutException("Le délai d'attente de la lecture sur le port série a expirée");
                }

                //Si on a déjà lu des données et qu'il n'y en a plus maintenant
                if (sBufferRead.Length != 0 && iNbData == 0)
                {
                    bStop = true;
                }
                Application.DoEvents();
            }
            while (!bStop);
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }

#if DEBUG
        //m_LogFile.WriteLine("ReadSerial - !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
        //m_LogFile.WriteLine(sBufferRead);
        //m_LogFile.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
#endif
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> private Int32 ReadSerial(Double dTimeOut, String sPrefix, String sSuffix,
    ///             out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode,
    ///             out String sErrorMessage, Boolean bStopReadWhenSuffixFind)</signature>
    ///
    /// <summary> Reads a serial.</summary>
    ///
    /// <param name="dTimeOut" type="Double">                 The time out. </param>
    /// <param name="sPrefix" type="String">                  The prefix. </param>
    /// <param name="sSuffix" type="String">                  The suffix. </param>
    /// <param name="sBufferRead" type="out String">          [in,out] The buffer read. </param>
    /// <param name="bErrorOccured" type="out Boolean">       [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">            [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">        [out] Message describing the error. </param>
    /// <param name="bStopReadWhenSuffixFind" type="Boolean"> true to stop read when suffix find. </param>
    ///
    /// <exception cref="TimeoutException"> Thrown when a Timeout error condition occurs. </exception>
    ///
    /// <returns> The serial.</returns>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    private Int32 ReadSerial(Double dTimeOut, String sPrefix, String sSuffix, out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage, Boolean bStopReadWhenSuffixFind)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;
        sBufferRead = String.Empty;

        try
        {
            Boolean bSearchPrefix = !String.IsNullOrEmpty(sPrefix);
            Boolean bSearchSuffix = !String.IsNullOrEmpty(sSuffix);
            Boolean bPrefixFound = false;
            Boolean bSuffixFound = false;

            int iSizeBuffer = m_ConfigSerial.iReadBufferSize;
            char[] arr = new char[iSizeBuffer];
            sBufferRead = "";

            bool bStop = false;
            int iNbData = 0;

            TimeSpan ts = TimeSpan.MinValue;
            DateTime dtStart = DateTime.Now;

            do
            {
                try
                {
                    iNbData = m_Serial.Read(arr, 0, iSizeBuffer - 1);
                    for (int i = 0; i < iNbData; i++)
                    {
                        sBufferRead += arr[i];
                    }

                    //si on attend un préfixe que l'on a pas encore trouvé
                    if (bSearchPrefix && !bPrefixFound)
                    {
                        //regarde la présence du préfixe dans la chaine reçue
                        bPrefixFound = sBufferRead.Contains(sPrefix);
                    }

                    //si on attend un suffixe que l'on a pas encore trouvé
                    if (bSearchSuffix && !bSuffixFound)
                    {
                        //si pas de prefixe à chercher
                        if (!bSearchPrefix)
                        {
                            //regarde la présence du suffixe dans la chaine reçue
                            bSuffixFound = sBufferRead.Contains(sSuffix);
                        }
                        else
                        {
                            //si le préfixe a été trouvé
                            if (bPrefixFound)
                            {
                                //regarde la présence du suffixe dans la chaine reçue à partir de l'emplacement du préfixe
                                int iIndexPrefixe = sBufferRead.IndexOf(sPrefix);
                                if (sBufferRead.Substring(iIndexPrefixe).Contains(sSuffix))
                                {
                                    bSuffixFound = true;
                                }
                            }
                        }
                    }
                }
                catch (TimeoutException)
                {
                    iNbData = 0;
                }

                //Gestion timeout
                ts = DateTime.Now - dtStart;
                if (ts.TotalMilliseconds > dTimeOut)
                {
                    throw new TimeoutException("Le délai d'attente de la lecture sur le port série a expirée");
                }


                //Si déjà lu des données
                if (sBufferRead.Length != 0)
                {
                    //Gestion des cas de sortie de la boucle en fonction de la recherche des prefixe et/ou suffixe

                    //PAS de recherche de prefixe ET PAS de recherche de suffixe
                    if (!bSearchPrefix && !bSearchSuffix)
                    {
                        if (iNbData == 0)
                            bStop = true;
                    }
                    //Recherche de prefixe ET Pas de recherche de suffixe
                    else if (bSearchPrefix && !bSearchSuffix)
                    {
                        if (iNbData == 0 && bPrefixFound)
                            bStop = true;
                    }
                    //PAS de recherche de prefixe ET Recherche de suffixe
                    else if (!bSearchPrefix && bSearchSuffix)
                    {
                        if (bSuffixFound && (bStopReadWhenSuffixFind || iNbData == 0))
                            bStop = true;
                    }
                    //Recherche de prefixe ET Recherche de suffixe
                    else
                    {
                        if (bPrefixFound && bSuffixFound && (bStopReadWhenSuffixFind || iNbData == 0))
                            bStop = true;
                    }
                }
                Application.DoEvents();
            }
            while (!bStop);

            bErrorOccured = false;
            iErrorCode = 0;
            sErrorMessage = "";
        }
        catch (Exception Ex)
        {
            bErrorOccured = true;
            iErrorCode = Ex.HResult;
            sErrorMessage = Ex.Message;
        }

#if DEBUG
        m_LogFile.WriteLine("ReadSerial {" + sPrefix.Replace("\r", "\\r").Replace("\n", "\\n") + "}{" + sSuffix.Replace("\r", "\\r").Replace("\n", "\\n") + "}" + "{" + iErrorCode.ToString() + "} - !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
        m_LogFile.WriteLine(sBufferRead);
        m_LogFile.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
#endif

        return iErrorCode;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void WriteCom(String sCommand, out Boolean bErrorOccured, out Int32 iErrorCode,
    ///             out String sErrorMessage)</signature>
    ///
    /// <summary> Writes a com.</summary>
    ///
    /// <param name="sCommand" type="String">           The command. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void WriteCom(String sCommand, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;

        ClearSerial(out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;

        WriteSerial(sCommand, out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void QueryCom(String sCommand, Double dTimeout, out String sBufferRead,
    ///             out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)</signature>
    ///
    /// <summary> Queries a com.</summary>
    ///
    /// <param name="sCommand" type="String">           The command. </param>
    /// <param name="dTimeout" type="Double">           The timeout. </param>
    /// <param name="sBufferRead" type="out String">    [in,out] The buffer read. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void QueryCom(String sCommand, Double dTimeout, out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;
        sBufferRead = String.Empty;

        ClearSerial(out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;

        WriteSerial(sCommand, out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;

        ReadSerial(dTimeout, "", "\r\n", out sBufferRead, out bErrorOccured, out iErrorCode, out sErrorMessage, true);
        if (bErrorOccured)
            return;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////
    /// <signature> public void QueryComPrefixSuffix(String sCommand, Double dTimeout, String sPrefix,
    ///             String sSuffix, out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode,
    ///             out String sErrorMessage)</signature>
    ///
    /// <summary> Queries com prefix suffix.</summary>
    ///
    /// <param name="sCommand" type="String">           The command. </param>
    /// <param name="dTimeout" type="Double">           The timeout. </param>
    /// <param name="sPrefix" type="String">            The prefix. </param>
    /// <param name="sSuffix" type="String">            The suffix. </param>
    /// <param name="sBufferRead" type="out String">    [out] The buffer read. </param>
    /// <param name="bErrorOccured" type="out Boolean"> [out] Flag to indicate than an error occurred. </param>
    /// <param name="iErrorCode" type="out Int32">      [out] Zero-based index of the error code. </param>
    /// <param name="sErrorMessage" type="out String">  [out] Message describing the error. </param>
    ////////////////////////////////////////////////////////////////////////////////////////////////////
    public void QueryComPrefixSuffix(String sCommand, Double dTimeout, String sPrefix, String sSuffix, out String sBufferRead, out Boolean bErrorOccured, out Int32 iErrorCode, out String sErrorMessage)
    {
        bErrorOccured = false;
        iErrorCode = 0;
        sErrorMessage = String.Empty;
        sBufferRead = String.Empty;

        ClearSerial(out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;

        WriteSerial(sCommand, out bErrorOccured, out iErrorCode, out sErrorMessage);
        if (bErrorOccured)
            return;

        ReadSerial(dTimeout, sPrefix, sSuffix, out sBufferRead, out bErrorOccured, out iErrorCode, out sErrorMessage, true);
        if (bErrorOccured)
            return;
    }
}