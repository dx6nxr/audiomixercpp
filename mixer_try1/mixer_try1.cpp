#include <iostream>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <audioclient.h>
#include <Windows.h>
#include <comdef.h>
#include <propkey.h>
#include <propvarutil.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <vector>
#include <string>

using namespace std;

struct AudioDevice {
    wstring name;
    IAudioEndpointVolume* endpointVolume;
};

int readFromSerial() {
    // Open the serial port
    HANDLE hSerial = CreateFile(L"\\\\.\\COM5",
        GENERIC_READ | GENERIC_WRITE,
        0,
        NULL,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        NULL);

    if (hSerial == INVALID_HANDLE_VALUE) {
        std::cerr << "Error opening COM5" << std::endl;
        return -1;
    }

    // Set the serial port parameters
    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
    if (!GetCommState(hSerial, &dcbSerialParams)) {
        std::cerr << "Error getting state" << std::endl;
        CloseHandle(hSerial);
        return -1;
    }

    dcbSerialParams.BaudRate = CBR_115200; // Set baud rate
    dcbSerialParams.ByteSize = 8;         // Data size
    dcbSerialParams.StopBits = ONESTOPBIT; // Stop bits
    dcbSerialParams.Parity = NOPARITY;    // Parity

    if (!SetCommState(hSerial, &dcbSerialParams)) {
        std::cerr << "Error setting state" << std::endl;
        CloseHandle(hSerial);
        return -1;
    }

    // Set timeouts
    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 1;
    timeouts.ReadTotalTimeoutConstant = 50;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    timeouts.WriteTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;

    SetCommTimeouts(hSerial, &timeouts);

    // Read data from the serial port
    char buffer[256];
    DWORD bytesRead;
    if (ReadFile(hSerial, buffer, sizeof(buffer) - 1, &bytesRead, NULL)) {
        buffer[bytesRead] = '\0'; // Null-terminate the string
        std::cout << "Read from COM5: " << buffer << std::endl;
    }
    else {
        std::cerr << "Error reading from COM5" << std::endl;
    }

    // Close the serial port
    CloseHandle(hSerial);

    return int(buffer);
}

void ListAudioSessionOutputs(vector<AudioDevice>& audioDevices) {
    CoInitialize(nullptr);

    IMMDeviceEnumerator* deviceEnumerator = nullptr;
    HRESULT hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        nullptr,
        CLSCTX_INPROC_SERVER,
        __uuidof(IMMDeviceEnumerator),
        (void**)&deviceEnumerator
    );

    if (SUCCEEDED(hr)) {
        IMMDeviceCollection* deviceCollection = nullptr;
        hr = deviceEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &deviceCollection);

        if (SUCCEEDED(hr)) {
            UINT count;
            deviceCollection->GetCount(&count);
            cout << "Number of audio output devices: " << count << endl;

            for (UINT i = 0; i < count; ++i) {
                IMMDevice* device = nullptr;
                hr = deviceCollection->Item(i, &device);
                if (SUCCEEDED(hr)) {
                    IPropertyStore* propertyStore = nullptr;
                    hr = device->OpenPropertyStore(STGM_READ, &propertyStore);

                    if (SUCCEEDED(hr)) {
                        PROPVARIANT varName;
                        PropVariantInit(&varName);
                        hr = propertyStore->GetValue(PKEY_Device_FriendlyName, &varName);

                        if (SUCCEEDED(hr)) {
                            wcout << i << L") Device: " << varName.pwszVal << endl;

                            // Create an AudioDevice struct and store it
                            AudioDevice audioDevice;
                            audioDevice.name = varName.pwszVal;
                            audioDevice.endpointVolume = nullptr;

                            // Activate the IAudioEndpointVolume interface
                            hr = device->Activate(
                                __uuidof(IAudioEndpointVolume),
                                CLSCTX_INPROC_SERVER,
                                nullptr,
                                (void**)&audioDevice.endpointVolume
                            );

                            if (SUCCEEDED(hr)) {
                                audioDevices.push_back(audioDevice); // Add to the list
                            }
                            else {
                                cout << "Failed to activate endpoint volume." << endl;
                            }

                            PropVariantClear(&varName);
                        }
                        else {
                            cout << "Failed to retrieve the name" << endl;
                        }

                        propertyStore->Release();
                    }
                    else {
                        cout << "Failed to get properties" << endl;
                    }
                    device->Release();
                }
                else {
                    cout << "Failed to get device" << endl;
                }
            }
            deviceCollection->Release();
        }
        deviceEnumerator->Release();
    }

    CoUninitialize();
}

void ChangeVolumeForAllDevices(const vector<AudioDevice>& audioDevices) {
    for (const auto& audioDevice : audioDevices) {
        if (audioDevice.endpointVolume) {
            float volumeLevel = 0.0f;
            HRESULT hr = audioDevice.endpointVolume->GetMasterVolumeLevelScalar(&volumeLevel);
            if (SUCCEEDED(hr)) {
                wcout << "Current Volume Level for " << wstring(audioDevice.name) << ": " << volumeLevel * 100 << "%" << endl;

                // Prompt user for new volume level
                float newVolume;
                wcout << "Enter new volume level for " << wstring(audioDevice.name) << " (0.0 to 1.0): ";
                cin >> newVolume;

                // Clamp the volume level between 0.0 and 1.0
                if (newVolume < 0.0f) newVolume = 0.0f;
                if (newVolume > 1.0f) newVolume = 1.0f;

                // Set the new volume level
                hr = audioDevice.endpointVolume->SetMasterVolumeLevelScalar(newVolume, nullptr);
                if (SUCCEEDED(hr)) {
                    cout << "Volume changed to: " << newVolume * 100 << "%" << endl;
                }

                else {
                        cout << "Failed to set the volume level." << endl;
                }
            }
            else {
                cout << "Failed to get the current volume level." << endl;
            }
        }
    }
}

void changeVolumeForChosenDevice(const vector<AudioDevice>& audioDevices, HANDLE hSerial) {
    cout << "Input device number:" << endl;
    int num;
    cin >> num;

    if (num >= 0 && num < audioDevices.size()) {
        const AudioDevice audioDevice = audioDevices[num];
        if (audioDevice.endpointVolume) {
            float volumeLevel = 0.0f;
            HRESULT hr = audioDevice.endpointVolume->GetMasterVolumeLevelScalar(&volumeLevel);
            // Read data from the serial port
            char buffer[256];
            DWORD bytesRead;
			float prevVolume = 0;
            while (true) {
                if (ReadFile(hSerial, buffer, sizeof(buffer) - 1, &bytesRead, NULL)) {
                    buffer[bytesRead] = '\0'; // Null-terminate the string
                    //std::cout << "Read from COM5: " << buffer << std::endl;
					// get the volume level from the serial port that should be between 0 and 1 float
					volumeLevel = atof(buffer) / 1023;
                    //cout << "volume level: " << volumeLevel << endl;
                    if (volumeLevel < 0.0f) volumeLevel = 0.0f;
                    if (volumeLevel > 1.0f) volumeLevel = 1.0f;

					// if the volume level has changed
					if (abs(volumeLevel - prevVolume) > 0.001) {
						// Set the new volume level
                        cout << "volume level: " << volumeLevel << endl;
                        hr = audioDevice.endpointVolume->SetMasterVolumeLevelScalar(volumeLevel, nullptr);
					}
                    
                    //cout << endl;
					//if (SUCCEEDED(hr)) {
					//	cout << "Volume changed to: " << volumeLevel * 100 << "%" << endl;
					//}
					//else {
					//	cout << "Failed to set the volume level." << endl;
					//}
                    // sleep 100 ms
					prevVolume = volumeLevel;
					Sleep(100);
                }
                else {
                    std::cerr << "Error reading from COM5" << std::endl;
                }
            }
            //cout << "setting the volume level to " << volumeLevel << endl;
        }
    }
}

int main() {
    vector<AudioDevice> audioDevices;

    // List audio session outputs and store them in the vector
    ListAudioSessionOutputs(audioDevices);

    //readFromSerial();

    HANDLE hSerial = CreateFile(L"\\\\.\\COM5",
        GENERIC_READ | GENERIC_WRITE,
        0,
        NULL,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        NULL);

    if (hSerial == INVALID_HANDLE_VALUE) {
        std::cerr << "Error opening COM5" << std::endl;
        return -1;
    }

    // Set the serial port parameters
    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
    if (!GetCommState(hSerial, &dcbSerialParams)) {
        std::cerr << "Error getting state" << std::endl;
        CloseHandle(hSerial);
        return -1;
    }

    dcbSerialParams.BaudRate = CBR_115200; // Set baud rate
    dcbSerialParams.ByteSize = 8;         // Data size
    dcbSerialParams.StopBits = ONESTOPBIT; // Stop bits
    dcbSerialParams.Parity = NOPARITY;    // Parity

    if (!SetCommState(hSerial, &dcbSerialParams)) {
        std::cerr << "Error setting state" << std::endl;
        CloseHandle(hSerial);
        return -1;
    }

    // Set timeouts
    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 1;
    timeouts.ReadTotalTimeoutConstant = 1;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    timeouts.WriteTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;

    SetCommTimeouts(hSerial, &timeouts);


    changeVolumeForChosenDevice(audioDevices, hSerial);

    // Change volume for all devices
    //ChangeVolumeForAllDevices(audioDevices);

    // Release the endpoint volume interfaces
    for (auto& audioDevice : audioDevices) {
        if (audioDevice.endpointVolume) {
            audioDevice.endpointVolume->Release();
        }
    }

    return 0;
}

