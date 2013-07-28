#pragma once

namespace FlashCap {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::IO::Compression;
	using namespace System::Runtime::InteropServices;

	/// <summary>
	/// Summary for CaptureForm
	/// </summary>
	[ComVisibleAttribute(true)]
	public ref class CaptureForm : public System::Windows::Forms::Form
	{
	public:
		CaptureForm(void)
		{
			this->WindowState = System::Windows::Forms::FormWindowState::Maximized;
			this->Load += gcnew System::EventHandler(this, &CaptureForm::CaptureForm_Load);
			ReadConfig();
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~CaptureForm()
		{
		}

	private: System::Windows::Forms::WebBrowser^  webBrowser;
	private: int timeout;
	private: array<unsigned int>^ framesToCapture;
	private: int framesToCaptureIndex;
	private: array<String^>^ files;
	private: int currentFileIndex;
	private: unsigned int currentFrame;
	private: String^ outputPath;
	private: String^ basePath;
	private: String^ webPath;
	private: String^ capturePath;
	private: String^ webBaseUrl;
	private: Timer^ timeoutTimer;

	private: System::Void ReadConfig() {
			this->basePath = AppDomain::CurrentDomain->BaseDirectory;
			this->webPath = Path::Combine(basePath, Configuration::ConfigurationManager::AppSettings["webPath"]) +
				Path::DirectorySeparatorChar;
			this->capturePath = Path::Combine(basePath, Configuration::ConfigurationManager::AppSettings["capturePath"]) +
				Path::DirectorySeparatorChar;
			if (Environment::GetCommandLineArgs()->Length > 2)
			{
				this->capturePath = Environment::GetCommandLineArgs()[2];
			}

			this->webBaseUrl = Configuration::ConfigurationManager::AppSettings["webUrl"];

			// Ensure directories exist
			if (!Directory::Exists(webPath))
			{
				Directory::CreateDirectory(webPath);
			}
			if (!Directory::Exists(capturePath))
			{
				Directory::CreateDirectory(capturePath);
			}

			timeout = UInt32::Parse(Configuration::ConfigurationManager::AppSettings["timeout"]);
			array<String^>^ framesParam = Configuration::ConfigurationManager::AppSettings["frames"]->Split(',');
			framesToCapture = gcnew array<unsigned int>(framesParam->Length);
			for (int i = 0; i < framesParam->Length; i++) {
				framesToCapture[i] = UInt32::Parse(framesParam[i]);
			}
			Array::Sort(framesToCapture);

			if (Environment::GetCommandLineArgs()->Length < 2 ||
				!Directory::Exists(Environment::GetCommandLineArgs()[1]))
			{
				throw gcnew Exception("Directory with SWF files is not specified");
			}
			files = Directory::GetFiles(Environment::GetCommandLineArgs()[1], "*.swf");
			if (files->Length == 0)
			{
				throw gcnew Exception("No SWF files found");
			}
			currentFileIndex = 0;
		}
	private: System::Void CaptureForm_Load(System::Object^  sender, System::EventArgs^  e) {
				 StartJob();
			 }
	private: System::Void webBrowser_DocumentCompleted(System::Object^  sender, System::Windows::Forms::WebBrowserDocumentCompletedEventArgs^  e) {
				 try {
					 HtmlElement^ element = webBrowser->Document->GetElementsByTagName("object")[0];
					 Rectangle offset = element->OffsetRectangle;
					 HtmlElement^ parent = element->OffsetParent;
					 while (parent) {
						 offset.Offset(parent->OffsetRectangle.Location);
						 parent  = parent->OffsetParent;
					 }
					 webBrowser->Size = offset.Size;
					 webBrowser->Document->Window->ScrollTo(offset.Location);
					 webBrowser->ObjectForScripting = this;
				 } catch (Exception^ ex) {
					 WriteMessage("Failed to locate the object: " + ex);
					 StopCapture();
				 }
			 }

	// public function for script for callback enterFrame notification
	public: System::Void OnEnterFrame() {
				try {
					if (framesToCapture[framesToCaptureIndex] == currentFrame) {
						Rectangle bounds = this->RectangleToScreen(webBrowser->Bounds);
						Bitmap^ bitmap = gcnew Bitmap(bounds.Width, bounds.Height);
						try
						{
							Graphics^ g = Graphics::FromImage(bitmap);
							try
							{
								g->CopyFromScreen(bounds.Location, Point::Empty, bounds.Size);
							} finally {
								delete g;
							}
							String^ pngPath = Path::Combine(outputPath, currentFrame + ".png");
							WriteMessage("Writting " + pngPath);
							bitmap->Save(pngPath, Imaging::ImageFormat::Png);
						} finally {
							delete bitmap;
						}
					}

					while (framesToCaptureIndex < framesToCapture->Length &&
						   currentFrame >= framesToCapture[framesToCaptureIndex]) {
						framesToCaptureIndex++;
					}

					currentFrame++;

					if (framesToCaptureIndex >= framesToCapture->Length) {
						timeoutTimer->Enabled = false;
						this->BeginInvoke(gcnew Action(this, &CaptureForm::StopCapture));
					}
				} catch (Exception^ ex) {
					WriteMessage("Failed to capture: " + ex);
					this->BeginInvoke(gcnew Action(this, &CaptureForm::StopCapture));
				}
			 }

	private: System::Void StartJob() {
			String^ swfPath = files[currentFileIndex];

			String^ jobName = Path::GetFileNameWithoutExtension(swfPath);
			this->Text = String::Format("Capturing {0} ({1}/{2})",
				jobName, currentFileIndex + 1, files->Length);

			outputPath = Path::Combine(capturePath, jobName);
			if (!Directory::Exists(outputPath))
			{
				Directory::CreateDirectory(outputPath);
			}

			WriteMessage("Capturing " + swfPath);
			try
			{
				framesToCaptureIndex = 0;
				StartCapture(swfPath);
			}
			catch (Exception^ ex)
			{
				WriteMessage("Capture failed: " + ex);
				NextJob();
			}
		}
	private: System::Void WriteMessage(String^ message) {
				 TextWriter^ writer = File::AppendText(Path::Combine(outputPath, "trace.log"));
				 writer->WriteLine(message);
				 writer->Close();
			 }
	private: System::Void NextJob() {
				 currentFileIndex++;
				 if (currentFileIndex < files->Length)
				 {
					 StartJob();
				 }
				 else
				 {
					 this->Close();
				 }
			 }

	private: System::Void StopCapture() {
				 webBrowser->DocumentCompleted -= gcnew System::Windows::Forms::WebBrowserDocumentCompletedEventHandler(this, &CaptureForm::webBrowser_DocumentCompleted);
				 webBrowser->ObjectForScripting = nullptr;
				 this->Controls->Remove(webBrowser);
				 delete [] webBrowser;
				 webBrowser = nullptr;
				 timeoutTimer->Tick -= gcnew System::EventHandler(this, &CaptureForm::timeoutTimer_Tick);
				 timeoutTimer->Enabled = false;
				 delete timeoutTimer;

				 NextJob();
			 }
	private: System::Void timeoutTimer_Tick(System::Object^  sender, System::EventArgs^  e) {
				 int timerFileIndex = (int)(((Timer^)sender)->Tag);
				 if (timerFileIndex == currentFileIndex) {
					 StopCapture();
				 }
			 }

	private: System::Void StartCapture(String^ swfPath) {
				 String^ jobId = Guid::NewGuid().ToString("D");
				 array<Byte>^ data = File::ReadAllBytes(swfPath);
				 if (data[0] == 'C' && data[1] == 'W' && data[2] == 'S')
				 {
					 // decompressing the data
					 WriteMessage("Decompressing");
					 unsigned int totalLength = BitConverter::ToUInt32(data, 4);
					 array<Byte>^ uncompressedData = gcnew array<Byte>(totalLength);
					 uncompressedData[0] = 'F';
					 uncompressedData[1] = 'W';
					 uncompressedData[2] = 'S';
					 uncompressedData[3] = data[3];
					 BitConverter::GetBytes(totalLength)->CopyTo(uncompressedData, 4);
					 MemoryStream^ dataStream = gcnew MemoryStream(data);
					 dataStream->Position = 8 + 2; // discarding swf header and zlib header
					 DeflateStream^ inflatedStream = gcnew DeflateStream(dataStream, CompressionMode::Decompress);
					 inflatedStream->Read(uncompressedData, 8, totalLength - 8);
					 delete inflatedStream;
					 delete dataStream;
					 data = uncompressedData;
				 }
				 if (data[0] != 'F' || data[1] != 'W' || data[2] != 'S')
				 {
					 throw gcnew Exception("Invalid SWF file header");
				 }
				 if (data[3] < 8)
				 {
					 throw gcnew Exception("The SWF file version is too old");
				 }
				 // finding the place to add frame tracer
				 unsigned int i = 8;
				 unsigned int frameSizePrecision = data[i] >> 3, frameSizeLength = (frameSizePrecision * 4 + 5 + 7) >> 3;
				 int frameSize[4];
				 for (unsigned j = 5, q = 0; q < 4; q++)
				 {
					 int n = (data[i + (j >> 3)] & (0x80 >> (j & 7))) ? -1 : 1;
					 j++;
					 for (unsigned p = 1; p < frameSizePrecision; p++, j++)
					 {
						n = (n << 1) | ((data[i + (j >> 3)] & (0x80 >> (j & 7))) ? 1 : 0);
					 }
					 frameSize[q] = n;
				 }
				 int width = (frameSize[1] - frameSize[0]) / 20, height = (frameSize[3] - frameSize[2]) / 20;
				 i += frameSizeLength;
				 WriteMessage("Size: " + width + "x" + height);

				 float frameRate = BitConverter::ToUInt16(data, i) / 256.0f;
				 i += 2;
				 WriteMessage("Frame rate: " + frameRate);

				 unsigned short frameCount = BitConverter::ToUInt16(data, i);
				 i += 2;
				 WriteMessage("Frame count: " + frameCount);

				 int tag = ((data[i + 1] << 2) | (data[i] >> 6));
				 if (tag != 69) // FileAttributes is required
				 {
					 throw gcnew Exception("No FileAttributes");
				 }
				 unsigned char fileAttributes = data[i + (((data[i] & 0x3F) == 0x3F) ? 6 : 2)];
				 bool isAvm2 = !!(fileAttributes & 0x08);
				 WriteMessage("AVM2: " + isAvm2);

				 unsigned short maxDepth = 0;
				 while (tag != 0 && tag != 1 && tag != 12) // skipping until End, ShowFrame or DoAction is found
				 {
					 unsigned int tagLength = data[i] & 0x3F, dataOffset = 2;
					 if (tagLength == 0x3F)
					 {
						 tagLength = BitConverter::ToUInt32(data, i + dataOffset);
						 dataOffset += 4;
					 }
					 switch (tag)
					 {
					 case 4: // PlaceObject
					 case 70: // PlaceObject3
						 maxDepth = Math::Max(maxDepth, BitConverter::ToUInt16(data, i + dataOffset + 2));
						 break;
					 case 26: // PlaceObject2
						 maxDepth = Math::Max(maxDepth, BitConverter::ToUInt16(data, i + dataOffset + 1));
						 break;
					 }

					 i += dataOffset + tagLength;
					 tag = ((data[i + 1] << 2) | (data[i] >> 6));
				 }

				 array<Byte>^ tracer;
				 if (isAvm2)
				 {
					 // see FlashControl/avm2.swf (was modified to make symbolId = 0xEFCD and depth to insert "nextDepth" bytes)
					 array<Byte>^ nextDepth = BitConverter::GetBytes(maxDepth + 1);
					 tracer = gcnew array<Byte> {0xFF, 0x09, 0x08, 0x00, 0x00, 0x00, 0xCD, 0xEF, 0x01, 0x00, 0x40, 0x00, 0x00, 0x00, 0x86, 0x06, 0x06,
						 nextDepth[0], nextDepth[1],
						 0xCD, 0xEF, 0x00, 0xBF, 0x14, 0xF7, 0x01, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x2E, 0x00, 0x00, 0x00, 0x00,
						 0x17, 0x0C, 0x66, 0x6C, 0x61, 0x73, 0x68, 0x2E, 0x65, 0x76, 0x65, 0x6E, 0x74, 0x73, 0x05, 0x45, 0x76, 0x65, 0x6E, 0x74, 0x00,
						 0x0D, 0x5F, 0x46, 0x72, 0x61, 0x6D, 0x65, 0x54, 0x72, 0x61, 0x63, 0x6B, 0x65, 0x72, 0x0D, 0x66, 0x6C, 0x61, 0x73, 0x68, 0x2E,
						 0x64, 0x69, 0x73, 0x70, 0x6C, 0x61, 0x79, 0x09, 0x4D, 0x6F, 0x76, 0x69, 0x65, 0x43, 0x6C, 0x69, 0x70, 0x0B, 0x69, 0x6E, 0x69,
						 0x74, 0x69, 0x61, 0x6C, 0x69, 0x7A, 0x65, 0x64, 0x07, 0x42, 0x6F, 0x6F, 0x6C, 0x65, 0x61, 0x6E, 0x0C, 0x6F, 0x6E, 0x45, 0x6E,
						 0x74, 0x65, 0x72, 0x46, 0x72, 0x61, 0x6D, 0x65, 0x0E, 0x66, 0x6C, 0x61, 0x73, 0x68, 0x2E, 0x65, 0x78, 0x74, 0x65, 0x72, 0x6E,
						 0x61, 0x6C, 0x11, 0x45, 0x78, 0x74, 0x65, 0x72, 0x6E, 0x61, 0x6C, 0x49, 0x6E, 0x74, 0x65, 0x72, 0x66, 0x61, 0x63, 0x65, 0x04,
						 0x63, 0x61, 0x6C, 0x6C, 0x05, 0x73, 0x74, 0x61, 0x67, 0x65, 0x04, 0x72, 0x6F, 0x6F, 0x74, 0x0B, 0x45, 0x4E, 0x54, 0x45, 0x52,
						 0x5F, 0x46, 0x52, 0x41, 0x4D, 0x45, 0x10, 0x61, 0x64, 0x64, 0x45, 0x76, 0x65, 0x6E, 0x74, 0x4C, 0x69, 0x73, 0x74, 0x65, 0x6E,
						 0x65, 0x72, 0x06, 0x4F, 0x62, 0x6A, 0x65, 0x63, 0x74, 0x0F, 0x45, 0x76, 0x65, 0x6E, 0x74, 0x44, 0x69, 0x73, 0x70, 0x61, 0x74,
						 0x63, 0x68, 0x65, 0x72, 0x0D, 0x44, 0x69, 0x73, 0x70, 0x6C, 0x61, 0x79, 0x4F, 0x62, 0x6A, 0x65, 0x63, 0x74, 0x11, 0x49, 0x6E,
						 0x74, 0x65, 0x72, 0x61, 0x63, 0x74, 0x69, 0x76, 0x65, 0x4F, 0x62, 0x6A, 0x65, 0x63, 0x74, 0x16, 0x44, 0x69, 0x73, 0x70, 0x6C,
						 0x61, 0x79, 0x4F, 0x62, 0x6A, 0x65, 0x63, 0x74, 0x43, 0x6F, 0x6E, 0x74, 0x61, 0x69, 0x6E, 0x65, 0x72, 0x06, 0x53, 0x70, 0x72,
						 0x69, 0x74, 0x65, 0x07, 0x16, 0x01, 0x16, 0x03, 0x16, 0x05, 0x18, 0x04, 0x05, 0x00, 0x16, 0x0A, 0x00, 0x13, 0x07, 0x01, 0x02,
						 0x07, 0x02, 0x04, 0x07, 0x03, 0x06, 0x07, 0x05, 0x07, 0x07, 0x02, 0x08, 0x07, 0x05, 0x09, 0x07, 0x06, 0x0B, 0x07, 0x02, 0x0C,
						 0x07, 0x02, 0x0D, 0x07, 0x02, 0x0E, 0x07, 0x02, 0x0F, 0x07, 0x02, 0x10, 0x07, 0x02, 0x11, 0x07, 0x01, 0x12, 0x07, 0x03, 0x13,
						 0x07, 0x03, 0x14, 0x07, 0x03, 0x15, 0x07, 0x03, 0x16, 0x04, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
						 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x09, 0x04, 0x00, 0x02, 0x00, 0x00, 0x02, 0x04, 0x00, 0x01, 0x05,
						 0x0A, 0x0A, 0x06, 0x11, 0x03, 0x01, 0x01, 0x03, 0x01, 0x02, 0x04, 0x01, 0x00, 0x04, 0x00, 0x02, 0x01, 0x09, 0x0A, 0x08, 0xD0,
						 0x30, 0x5E, 0x04, 0x27, 0x61, 0x04, 0x47, 0x00, 0x00, 0x01, 0x02, 0x02, 0x09, 0x0A, 0x0A, 0xD0, 0x30, 0x60, 0x07, 0x2C, 0x09,
						 0x4F, 0x08, 0x01, 0x47, 0x00, 0x00, 0x02, 0x04, 0x01, 0x0A, 0x0B, 0x1F, 0xD0, 0x30, 0xD0, 0x49, 0x00, 0x60, 0x04, 0x11, 0x13,
						 0x00, 0x00, 0x5E, 0x04, 0x26, 0x61, 0x04, 0x60, 0x09, 0x66, 0x0A, 0x60, 0x01, 0x66, 0x0B, 0x60, 0x06, 0x27, 0x4F, 0x0C, 0x03,
						 0x47, 0x00, 0x00, 0x03, 0x02, 0x01, 0x01, 0x09, 0x27, 0xD0, 0x30, 0x65, 0x00, 0x60, 0x0D, 0x30, 0x60, 0x0E, 0x30, 0x60, 0x0F,
						 0x30, 0x60, 0x10, 0x30, 0x60, 0x11, 0x30, 0x60, 0x12, 0x30, 0x60, 0x03, 0x30, 0x60, 0x03, 0x58, 0x00, 0x1D, 0x1D, 0x1D, 0x1D,
						 0x1D, 0x1D, 0x1D, 0x68, 0x02, 0x47, 0x00, 0x00, 0x3F, 0x13, 0x12, 0x00, 0x00, 0x00, 0x01, 0x00, 0xCD, 0xEF, 0x5F, 0x46, 0x72,
						 0x61, 0x6D, 0x65, 0x54, 0x72, 0x61, 0x63, 0x6B, 0x65, 0x72, 0x00};
				 }
				 else
				 {
					 // see FlashControl/avm1.swf
					 tracer = gcnew array<Byte> {0x3F, 0x03, 0xB8, 0x00, 0x00, 0x00, 0x88, 0x51, 0x00, 0x08, 0x00, 0x74, 0x68, 0x69, 0x73, 0x00, 0x07,
						 0x00, 0x63, 0x72, 0x65, 0x61, 0x74, 0x65, 0x45, 0x6D, 0x70, 0x74, 0x79, 0x4D, 0x6F, 0x76, 0x69, 0x65, 0x43, 0x6C, 0x69, 0x70,
						 0x00, 0x6F, 0x6E, 0x45, 0x6E, 0x74, 0x65, 0x72, 0x46, 0x72, 0x61, 0x6D, 0x65, 0x00, 0x66, 0x6C, 0x61, 0x73, 0x68, 0x00, 0x65,
						 0x78, 0x74, 0x65, 0x72, 0x6E, 0x61, 0x6C, 0x00, 0x45, 0x78, 0x74, 0x65, 0x72, 0x6E, 0x61, 0x6C, 0x49, 0x6E, 0x74, 0x65, 0x72,
						 0x66, 0x61, 0x63, 0x65, 0x00, 0x63, 0x61, 0x6C, 0x6C, 0x00, 0x96, 0x02, 0x00, 0x08, 0x00, 0x1C, 0x96, 0x02, 0x00, 0x08, 0x01,
						 0x4E, 0x4C, 0x9D, 0x02, 0x00, 0x50, 0x00, 0x17, 0x96, 0x0E, 0x00, 0x07, 0xE8, 0x03, 0x00, 0x00, 0x08, 0x01, 0x07, 0x02, 0x00,
						 0x00, 0x00, 0x08, 0x00, 0x1C, 0x96, 0x02, 0x00, 0x08, 0x02, 0x52, 0x96, 0x02, 0x00, 0x08, 0x03, 0x9B, 0x05, 0x00, 0x00, 0x00,
						 0x00, 0x20, 0x00, 0x96, 0x09, 0x00, 0x08, 0x03, 0x07, 0x01, 0x00, 0x00, 0x00, 0x08, 0x04, 0x1C, 0x96, 0x02, 0x00, 0x08, 0x05,
						 0x4E, 0x96, 0x02, 0x00, 0x08, 0x06, 0x4E, 0x96, 0x02, 0x00, 0x08, 0x07, 0x52, 0x17, 0x87, 0x01, 0x00, 0x00, 0x4F, 0x96, 0x02,
						 0x00, 0x04, 0x00, 0x17, 0x00};
				 }
				 array<Byte>^ newData = gcnew array<Byte>(tracer->Length + data->Length);
				 array<Byte>::Copy(data, 0, newData, 0, i);
				 array<Byte>::Copy(tracer, 0, newData, i, tracer->Length);
				 array<Byte>::Copy(data, i, newData, i + tracer->Length, data->Length - i);
				 BitConverter::GetBytes(UInt32(newData->Length))->CopyTo(newData, 4);

				 File::WriteAllBytes(webPath + jobId + ".swf", newData);

				 String^ wrapper = File::ReadAllText(basePath + "wrapper.html");
				 wrapper = wrapper->
					 Replace("movie.swf", jobId + ".swf")->
					 Replace("640", width.ToString())->
					 Replace("480", height.ToString());
				 File::WriteAllText(webPath + jobId + ".html", wrapper);

  				 this->currentFrame = isAvm2 ? 0 : 1; // avm1 makes all visible on first frame

				 webBrowser = gcnew WebBrowser();
			     webBrowser->ScrollBarsEnabled = false;
			     webBrowser->Size = System::Drawing::Size(100, 100);
			     webBrowser->DocumentCompleted += gcnew System::Windows::Forms::WebBrowserDocumentCompletedEventHandler(this, &CaptureForm::webBrowser_DocumentCompleted);
			     this->Controls->Add(webBrowser);

				 timeoutTimer = gcnew Timer();
				 timeoutTimer->Interval = timeout;
				 timeoutTimer->Enabled = true;
				 timeoutTimer->Tick += gcnew System::EventHandler(this, &CaptureForm::timeoutTimer_Tick);
				 timeoutTimer->Tag = currentFileIndex;

				 String^ url = webBaseUrl + jobId + ".html";
				 this->webBrowser->Url = gcnew System::Uri(url, System::UriKind::Absolute);
			 }
};
}

