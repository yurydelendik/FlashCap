// FlashCap.cpp : main project file.

#include "stdafx.h"
#include "CaptureForm.h"

using namespace FlashCap;

[STAThreadAttribute]
int main(array<System::String ^> ^args)
{
	// Create the main window and run it
	Application::Run(gcnew CaptureForm());
	return 0;
}
