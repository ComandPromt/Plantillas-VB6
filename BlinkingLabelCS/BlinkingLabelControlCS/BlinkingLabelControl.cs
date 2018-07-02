using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace BlinkingLabelControlCS
{
	/// <summary>
	/// Summary description for BlinkingLabelControl.
	/// </summary>
	/// 



	public class BlinkingLabelControl : System.Windows.Forms.Control
	{

		// Variables used internally
		
		private System.Timers.Timer myTimer = new System.Timers.Timer();
		private int m_BlinkInterval = 1 ;
		private System.Drawing.Color m_BlinkoffColor = System.Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control) ;
		private System.Drawing.SolidBrush m_BlinkOnBrush ;
		private System.Drawing.SolidBrush m_BlinkOffBrush ;
		private bool m_UseBlinkOnColor = true ;

		public BlinkingLabelControl()
		{
			// Hook up handler function for timer, and set timer variables

			myTimer.Elapsed += new System.Timers.ElapsedEventHandler (this.OnTimerExpired) ;
			myTimer.AutoReset = true ;
			myTimer.Interval = 1000 ;
			myTimer.Enabled = true ;

			// Create brushes used for drawing label text

			m_BlinkOnBrush = new System.Drawing.SolidBrush(ForeColor) ;
			m_BlinkOffBrush = new System.Drawing.SolidBrush(BackColor) ;

		}

		protected override void OnPaint(PaintEventArgs pe)
		{
			// Calling the base class OnPaint
			base.OnPaint(pe);

			System.Drawing.Brush BrushToUse ;

			// Choose the brush to use for the text color. If the blink cycle is
			// currently on, or if we're in design mode (in which case we never
			// want to blink, select the first brush. Otherwise select the second.

			if (m_UseBlinkOnColor == true || DesignMode == true) 
			{
				BrushToUse = m_BlinkOnBrush ;
			}
			else
			{
				BrushToUse = m_BlinkOffBrush ;
			}

			// Draw the control's current Text property

			pe.Graphics.DrawString(Text, Font, BrushToUse, 0, 0) ;

		}
		// Control properties. This is the interval at which the blinking label
		// changes its color from on to off or back again, in seconds.

		public int BlinkInterval
		{
			get
			{
				return m_BlinkInterval ;
			}
			set
			{
				m_BlinkInterval = value ;
				myTimer.Interval = m_BlinkInterval * 1000 ;
			}
		}

		// This property is the color in which the blinking label's text is shown
		// during the Off part of the blink cycle. This allows you to make the blinking
		// label appear to turn on and off, or to switch from one color to another.
		// The control's ForeColor property is always used as the BlinkOn color.

		public System.Drawing.Color BlinkOffColor
		{
			get
			{
				return m_BlinkoffColor ;
			}
			set
			{
				m_BlinkoffColor = value ;
				m_BlinkOffBrush = new System.Drawing.SolidBrush(m_BlinkoffColor) ;
			}
		}

		// When the ForeColor property, which we inherit from the base class, changes,
		// then create a new brush with the new ForeColor to use in painting the label text.


		protected override void OnForeColorChanged(System.EventArgs e)
		{
			m_BlinkOnBrush = new System.Drawing.SolidBrush (ForeColor) ;
		}

		//  Handler for our control's internal timer. 

		private void OnTimerExpired( Object Source,  System.Timers.ElapsedEventArgs e)
		{

			// Toggle the flag the tells the painting code whether to use the BlinkOnColor
			//  or the BlinkOff color

			if (m_UseBlinkOnColor == true)
			{
				m_UseBlinkOnColor = false ;
			}
			else
			{
				m_UseBlinkOnColor = true ;
			}
			
			// Invalidate the control to force a repaint.

			Invalidate() ;

			// Fire the blink event to the control's container, in case it cares.

			BlinkStateChanged (m_UseBlinkOnColor) ;
		}

		// Declare an event that this control will fire to its container. This
		// event is called Blink, and is fired when the internal timer expires
		// and the blinking label changes color. The single parameter tells the
		// event recipient whether the color is changing to BlinkOn (true) or to
		// BlinkOff (false).

		public delegate void BlinkStateChangedHandler (bool UseBlinkColor) ;
		public event  BlinkStateChangedHandler BlinkStateChanged ;
	}
}
