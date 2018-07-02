#pragma once

using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;


namespace WindowsApp
{
   // Richard Grimes:
   // Ignore this warning. If you have followed the steps outlined in
   // the ReadMe.txt file you will have a form with a meaningful name.

	/// <summary> 
	/// Summary for MainForm
	///
	/// WARNING: If you change the name of this class, you will need to change the 
	///          'Resource File Name' property for the managed resource compiler tool 
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public __gc class MainForm : public System::Windows::Forms::Form
	{
	public: 
		MainForm(void)
		{
			InitializeComponent();
		}
        
	protected: 
		void Dispose(Boolean disposing)
		{
			if (disposing && components)
			{
				components->Dispose();
			}
			__super::Dispose(disposing);
		}
   private: System::Windows::Forms::Label *  label1;

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container* components;

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
         this->label1 = new System::Windows::Forms::Label();
         this->SuspendLayout();
         // 
         // label1
         // 
         this->label1->Location = System::Drawing::Point(8, 8);
         this->label1->Name = S"label1";
         this->label1->Size = System::Drawing::Size(280, 40);
         this->label1->TabIndex = 0;
         this->label1->Text = S"This form does nothing, it is just an example of how to clean up the wizard gener" 
            S"ated code";
         // 
         // MainForm
         // 
         this->AutoScaleBaseSize = System::Drawing::Size(6, 15);
         this->ClientSize = System::Drawing::Size(292, 48);
         this->Controls->Add(this->label1);
         this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedDialog;
         this->Name = S"MainForm";
         this->Text = S"MainForm";
         this->ResumeLayout(false);

      }		
	};
}