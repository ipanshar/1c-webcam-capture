using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Drawing;
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
    [ProgId("WebCameraCapture. WebCameraCapture")]
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
        private byte[] lastFrameData;

        public WebCameraCapture()
        {
            lastFrameData = new byte[0];
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

                // Get the capture device
                Guid iid = typeof(IBaseFilter).GUID;
                object obj;
                videoCaptureDevices[selectedCameraIndex].Mon.BindToObject(null, null, ref iid, out obj);
                captureFilter = (IBaseFilter)obj;

                // Add capture filter to graph
                filterGraph.AddFilter(captureFilter, "Video Capture");

                // Create and add sample grabber
                sampleGrabberFilter = (IBaseFilter)new SampleGrabber();
                sampleGrabber = (ISampleGrabber)sampleGrabberFilter;

                // Configure sample grabber
                AMMediaType media = new AMMediaType();
                media.majorType = MediaType.Video;
                media.subType = MediaSubType.RGB24;
                sampleGrabber.SetMediaType(media);

                filterGraph.AddFilter(sampleGrabberFilter, "Sample Grabber");

                // Connect filters
                IPin outPin = DsFindPin.ByDirection(captureFilter, PinDirection.Output, 0);
                IPin inPin = DsFindPin.ByDirection(sampleGrabberFilter, PinDirection.Input, 0);
                filterGraph.Connect(outPin, inPin);

                // Start preview
                mediaControl.Run();

                lastError = "";
                return true;
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return false;
            }
        }

        public byte[] GetLastFrameData()
        {
            return lastFrameData;
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

                // First, get the size of the buffer
                int hr = sampleGrabber.GetCurrentBuffer(ref bufferSize, IntPtr.Zero);
                if (hr != 0 || bufferSize <= 0)
                {
                    lastError = "Failed to get buffer size";
                    return new byte[0];
                }

                // Allocate the buffer and retrieve the frame data
                IntPtr bufferPtr = Marshal.AllocCoTaskMem(bufferSize);
                try
                {
                    hr = sampleGrabber.GetCurrentBuffer(ref bufferSize, bufferPtr);
                    if (hr != 0)
                    {
                        lastError = "Failed to capture frame";
                        return new byte[0];
                    }

                    byte[] frameData = new byte[bufferSize];
                    Marshal.Copy(bufferPtr, frameData, 0, bufferSize);

                    // Convert to JPEG
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

        public string[] GetDeviceList()
        {
            try
            {
                List<string> devices = new List<string>();

                foreach (DsDevice dev in videoCaptureDevices)
                {
                    devices.Add(dev.Name);
                }

                lastError = "";
                return devices.ToArray();
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return new string[0];
            }
        }

        public bool SetCamera(int index)
        {
            try
            {
                if (index < 0 || index >= videoCaptureDevices.Length)
                {
                    lastError = "Invalid camera index";
                    return false;
                }

                selectedCameraIndex = index;
                Cleanup();
                Initialize();

                lastError = "";
                return true;
            }
            catch (Exception ex)
            {
                lastError = ex.Message;
                return false;
            }
        }

        public string GetErrorMessage()
        {
            return lastError;
        }

        public int GetCameraCount()
        {
            return videoCaptureDevices.Length;
        }

        public void Cleanup()
        {
            try
            {
                if (mediaControl != null)
                {
                    mediaControl.Stop();
                }

                if (captureFilter != null)
                {
                    Marshal.ReleaseComObject(captureFilter);
                    captureFilter = null;
                }

                if (sampleGrabberFilter != null)
                {
                    Marshal.ReleaseComObject(sampleGrabberFilter);
                    sampleGrabberFilter = null;
                }

                if (filterGraph != null)
                {
                    Marshal.ReleaseComObject(filterGraph);
                    filterGraph = null;
                }

                sampleGrabber = null;
                mediaControl = null;
            }
            catch { }
        }

        private byte[] ConvertToJpeg(byte[] rgbData)
        {
            try
            {
                // Simplified JPEG conversion
                // In production, use a proper image processing library
                using (var ms = new MemoryStream())
                {
                    // Placeholder for JPEG encoding
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                lastError = "JPEG conversion failed: " + ex.Message;
                return new byte[0];
            }
        }
    }
}