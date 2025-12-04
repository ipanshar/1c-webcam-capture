using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid("YOUR_GUID_HERE")]  
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWebCameraCapture  
{
    void Initialize();
    void CaptureFrame();
    string GetDeviceList();
    void SetCamera(int index);
    string GetErrorMessage();
    void Cleanup();
    int GetCameraCount();
}

[ComVisible(true)]
[Guid("YOUR_COM_CLASS_GUID_HERE")]  
[ClassInterface(ClassInterfaceType.None)]
public class WebCameraCapture : IWebCameraCapture  
{
    public void Initialize()
    {
        // Implementation for initializing webcam
    }

    public void CaptureFrame()
    {
        // Implementation for capturing a frame from the webcam
    }

    public string GetDeviceList()
    {
        // Implementation for getting a list of webcam devices
        return string.Empty;
    }

    public void SetCamera(int index)
    {
        // Implementation for setting the webcam by index
    }

    public string GetErrorMessage()
    {
        // Implementation for getting the last error message
        return string.Empty;
    }

    public void Cleanup()
    {
        // Implementation for cleaning up resources
    }

    public int GetCameraCount()
    {
        // Implementation for getting the count of available webcams
        return 0;
    }
}