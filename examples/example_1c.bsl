// 1C Examples for Using the Webcam DLL

// Example 1: Initialize the Webcam
Webcam.Initialize();

// Example 2: Capture an Image
Image = Webcam.CaptureImage();

// Example 3: Save the Captured Image
File.Save(Image, "captured_image.png");

// Example 4: Release the Webcam
Webcam.Release();