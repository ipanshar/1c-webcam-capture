using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DirectShowLib;

namespace WebCameraCSharp
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789012")]
    public interface IWebCameraCapture
    {
        [DispId(1)]
        byte[] CaptureFrame();

        [DispId(2)]
        string[] GetDeviceList();

        [DispId(3)]
        bool SetCamera(int index);

        [DispId(4)]
        string GetErrorMessage();

        [DispId(5)]
        bool Initialize();

        [DispId(6)]
        void Cleanup();

        [DispId(7)]
        int GetCameraCount();
    }

    [ComVisible(true)]
    [Guid("87654321-4321-4321-4321-210987654321")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("WebCameraCapture.WebCameraCapture")]
    public class WebCameraCapture : IWebCameraCapture
    {
        private IFilterGraph2 filterGraph;
        private IMediaControl mediaControl;
        private ISampleGrabber sampleGrabber;
        private IBaseFilter captureFilter;
        private IBaseFilter sampleGrabberFilter;
        private int selectedCameraIndex = 0;
        private string lastError = "";
        private DsDevice[] videoCaptureDevices;
        private int videoWidth;
        private int videoHeight;

        public WebCameraCapture()
        {
            videoCaptureDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);
        }

        public bool Initialize()
        {
            try
            {
                if (videoCaptureDevices.Length == 0)
                {
                    lastError = "No video capture devices found";
                    return false;
                }

                filterGraph = (IFilterGraph2)new FilterGraph();
                mediaControl = (IMediaControl)filterGraph;

                // Получаем выбранную камеру
                Guid iid = typeof(IBaseFilter).GUID;
                object obj;
                videoCaptureDevices[selectedCameraIndex].Mon.BindToObject(null, null, ref iid, out obj);
                captureFilter = (IBaseFilter)obj;

                filterGraph.AddFilter(captureFilter, "Video Capture");

                // Создаем SampleGrabber
                sampleGrabberFilter = (IBaseFilter)new SampleGrabber();
                sampleGrabber = (ISampleGrabber)sampleGrabberFilter;

                // Настраиваем формат
                AMMediaType media = new AMMediaType
                {
                    majorType = MediaType.Video,
                    subType = MediaSubType.RGB24,
                    formatType = FormatType.VideoInfo
                };
                sampleGrabber.SetMediaType(media);

                sampleGrabber.SetBufferSamples(true);
                sampleGrabber.SetCallback(null, 0);

                filterGraph.AddFilter(sampleGrabberFilter, "Sample Grabber");

                // Соединяем фильтры
                IPin outPin = DsFindPin.ByDirection(captureFilter, PinDirection.Output, 0);
                IPin inPin = DsFindPin.ByDirection(sampleGrabberFilter, PinDirection.Input, 0);
                filterGraph.Connect(outPin, inPin);

                // Запускаем граф
                mediaControl.Run();

                // Получаем параметры кадра
                AMMediaType connectedMedia = new AMMediaType();
                sampleGrabber.GetConnectedMediaType(connectedMedia);
                if ((connectedMedia.formatType != FormatType.VideoInfo) || (connectedMedia.formatPtr == IntPtr.Zero))
                    throw new Exception("Unsupported media type");

                VideoInfoHeader vih = (VideoInfoHeader)Marshal.PtrToStructure(connectedMedia.formatPtr, typeof(VideoInfoHeader));
                videoWidth = vih.BmiHeader.Width;
                videoHeight = vih.BmiHeader.Height;

                DsUtils.FreeAMMediaType(connectedMedia);

                lastError = "";
                return true;
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return false;
            }
        }

        public byte[] CaptureFrame()
        {
            try
            {
                if (sampleGrabber == null)
                {
                    lastError = "Camera not initialized";
                    return new byte[0];
                }

                int bufferSize = 0;
                int hr = sampleGrabber.GetCurrentBuffer(ref bufferSize, IntPtr.Zero);
                if (hr != 0 || bufferSize <= 0)
                {
                    lastError = $"Failed to get buffer size, hr={hr}";
                    return new byte[0];
                }

                IntPtr bufferPtr = Marshal.AllocCoTaskMem(bufferSize);
                try
                {
                    hr = sampleGrabber.GetCurrentBuffer(ref bufferSize, bufferPtr);
                    if (hr != 0)
                    {
                        lastError = $"Failed to capture frame, hr={hr}";
                        return new byte[0];
                    }

                    byte[] frameData = new byte[bufferSize];
                    Marshal.Copy(bufferPtr, frameData, 0, bufferSize);

                    return ConvertToJpeg(frameData);
                }
                finally
                {
                    Marshal.FreeCoTaskMem(bufferPtr);
                }
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return new byte[0];
            }
        }

        private byte[] ConvertToJpeg(byte[] rgbData)
        {
            try
            {
                Bitmap bmp = new Bitmap(videoWidth, videoHeight, PixelFormat.Format24bppRgb);
                BitmapData bmpData = bmp.LockBits(new Rectangle(0, 0, videoWidth, videoHeight),
                    ImageLockMode.WriteOnly, bmp.PixelFormat);

                Marshal.Copy(rgbData, 0, bmpData.Scan0, rgbData.Length);
                bmp.UnlockBits(bmpData);

                using (MemoryStream ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Jpeg);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                lastError = "JPEG conversion failed: " + ex.Message;
                return new byte[0];
            }
        }

        public string[] GetDeviceList()
        {
            try
            {
                var devices = new string[videoCaptureDevices.Length];
                for (int i = 0; i < videoCaptureDevices.Length; i++)
                    devices[i] = videoCaptureDevices[i].Name;
                lastError = "";
                return devices;
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return new string[0];
            }
        }

        public bool SetCamera(int index)
        {
            if (index < 0 || index >= videoCaptureDevices.Length)
            {
                lastError = "Invalid camera index";
                return false;
            }
            selectedCameraIndex = index;
            Cleanup();
            return Initialize();
        }

        public string GetErrorMessage() => lastError;

        public int GetCameraCount() => videoCaptureDevices.Length;

        public void Cleanup()
        {
            try
            {
                if (mediaControl != null)
                    mediaControl.Stop();

                if (captureFilter != null)
                    Marshal.ReleaseComObject(captureFilter);
                if (sampleGrabberFilter != null)
                    Marshal.ReleaseComObject(sampleGrabberFilter);
                if (filterGraph != null)
                    Marshal.ReleaseComObject(filterGraph);

                captureFilter = null;
                sampleGrabberFilter = null;
                filterGraph = null;
                sampleGrabber = null;
                mediaControl = null;
            }
            catch { }
        }
    }
}
