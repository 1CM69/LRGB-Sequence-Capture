import clr
import time

clr.AddReference("SharpCap")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Threading")

from System.Threading import Thread, ThreadStart
from System.Windows.Forms import Application, Form, FormStartPosition, TextBox, Button, Label, MessageBox, MessageBoxButtons, MessageBoxIcon, ScrollBars, GroupBox, RadioButton, NumericUpDown
from System import Environment, TimeSpan
from SharpCap.UI import CaptureLimitType

# GLOSSARY OF OBJECTS
# gbL       -> gbB      : GroupBox      LUMINANCE -> BLUE
# lblLE     -> lblBE    : Label         LUMINANCE -> BLUE                            : EXPOSURE
# lblLEv    -> lblBEv   : Label         LUMINANCE -> BLUE                            : EXPOSURE VALUE
# btnLE     -> btnBE    : Button        LUMINANCE -> BLUE                            : EXPOSURE
# lblLG     -> lblBG    : Label         LUMINANCE -> BLUE                            : GAIN
# lblLGv    -> lblBGv   : Label         LUMINANCE -> BLUE                            : GAIN VALUE
# btnLG     -> btnBG    : Button        LUMINANCE -> BLUE                            : GAIN
# gbTFS                 : GroupBox      TIMELIMIT, FRAMELIMIT & SEQUENCE SECTION
# rbTFST                : RadioButton   -------------------""-------------------     : TIMELIMIT
# rbTFSF                : RadioButton   -------------------""-------------------     : FRAMELIMIT
# rbTF                  :               -------------------""-------------------     : EVENT USED TO DETERMINE WHICH RADIOBUTTON (above) HAS BEEN CHECKED
# lblTFSTF              : Label         -------------------""-------------------     : TIME/FRAME LIMIT
# nudTFSTF              : NumericUpDown -------------------""-------------------     : TIME/FRAME LIMIT
# lblTFSS               : Label         -------------------""-------------------     : NUMBER OF SEQUENCES
# nudTFSS               : NumericUpDown -------------------""-------------------     : NUMBER OF SEQUENCES
# gbProgress            : GroupBox      PROGRESS SECTION
# tbOutput              : TextBox       -------""-------                             : INFORMATION WINDOW
# btnSC                 : Button                                                     : START CAPTURE
# get_EV()              : Funtion                                                    : GET & FORMAT EXPOSURE VALUES FROM SHARPCAP                                                               

capTYPE = 0
SCVer = float(SharpCap.GetType().Assembly.GetName().Version.Major.ToString() + '.' + SharpCap.GetType().Assembly.GetName().Version.Minor.ToString())

class Main(Form):
    def __init__(self):
        Form.__init__(self)
        # MAIN FORM CONSTRUCTION
        self.Text = 'LRGB Sequence Capture'
        self.TopMost = True
        self.Width = 780
        self.Height = 460
        self.ShowIcon = False
        self.MaximizeBox = False
        
        # LUMINANCE FILTER SECTION
        self.gbL = GroupBox()
        self.gbL.Left = 25
        self.gbL.Top = 25
        self.gbL.Width = 160
        self.gbL.Height = 130
        self.gbL.Text = "LUMINANCE FILTER #1"

        self.lblLE = Label()
        self.lblLE.Left = 10
        self.lblLE.Top = 25
        self.lblLE.Text = "Exposure: "
        self.lblLE.AutoSize = True
        self.gbL.Controls.Add(self.lblLE)
        
        self.lblLEv = Label()
        self.lblLEv.Left = self.lblLE.Left + self.lblLE.Width
        self.lblLEv.Top = 25
        self.lblLEv.Text = ""
        self.lblLEv.AutoSize = True
        self.gbL.Controls.Add(self.lblLEv)

        self.btnLE = Button()
        self.btnLE.Left = 10
        self.btnLE.Top = 40
        self.btnLE.Width = 140
        self.btnLE.Text = "Get LUM Exposure"
        self.btnLE.Click += self.btnLE_Click
        self.gbL.Controls.Add(self.btnLE)

        self.lblLG = Label()
        self.lblLG.Left = 10
        self.lblLG.Top = 75
        self.lblLG.Text = "Gain: "
        self.lblLG.AutoSize = True
        self.gbL.Controls.Add(self.lblLG)
        
        self.lblLGv = Label()
        self.lblLGv.Left = self.lblLG.Left + self.lblLG.Width
        self.lblLGv.Top = 75
        self.lblLGv.Text = ""
        self.lblLGv.AutoSize = True
        self.gbL.Controls.Add(self.lblLGv)
        
        self.btnLG = Button()
        self.btnLG.Left = 10
        self.btnLG.Top = 90
        self.btnLG.Width = 140
        self.btnLG.Text = "Get LUM Gain"
        self.btnLG.Click += self.btnLG_Click
        self.gbL.Controls.Add(self.btnLG)

        self.Controls.Add(self.gbL)

        # RED FILTER SECTION
        self.gbR = GroupBox()
        self.gbR.Left = 210
        self.gbR.Top = 25
        self.gbR.Width = 160
        self.gbR.Height = 130
        self.gbR.Text = "RED FILTER #2"

        self.lblRE = Label()
        self.lblRE.Left = 10
        self.lblRE.Top = 25
        self.lblRE.Width = 140
        self.lblRE.Text = "Exposure: "
        self.lblRE.AutoSize = True
        self.gbR.Controls.Add(self.lblRE)

        self.lblREv = Label()
        self.lblREv.Left = self.lblRE.Left + self.lblRE.Width
        self.lblREv.Top = 25
        self.lblREv.Text = ""
        self.lblREv.AutoSize = True
        self.gbR.Controls.Add(self.lblREv)

        self.btnRE = Button()
        self.btnRE.Left = 10
        self.btnRE.Top = 40
        self.btnRE.Width = 140
        self.btnRE.Text = "Get RED Exposure"
        self.btnRE.Click += self.btnRE_Click
        self.gbR.Controls.Add(self.btnRE)

        self.lblRG = Label()
        self.lblRG.Left = 10
        self.lblRG.Top = 75
        self.lblRG.Text = "Gain: "
        self.lblRG.AutoSize = True
        self.gbR.Controls.Add(self.lblRG)

        self.lblRGv = Label()
        self.lblRGv.Left = self.lblRG.Left + self.lblRG.Width
        self.lblRGv.Top = 75
        self.lblRGv.Text = ""
        self.lblRGv.AutoSize = True
        self.gbR.Controls.Add(self.lblRGv)
        
        self.btnRG = Button()
        self.btnRG.Left = 10
        self.btnRG.Top = 90
        self.btnRG.Width = 140
        self.btnRG.Text = "Get RED Gain"
        self.btnRG.Click += self.btnRG_Click
        self.gbR.Controls.Add(self.btnRG)

        self.Controls.Add(self.gbR)

        # GREEN FILTER SECTION
        self.gbG = GroupBox()
        self.gbG.Left = 395
        self.gbG.Top = 25
        self.gbG.Width = 160
        self.gbG.Height = 130
        self.gbG.Text = "GREEN FILTER #3"

        self.lblGE = Label()
        self.lblGE.Left = 10
        self.lblGE.Top = 25
        self.lblGE.Text = "Exposure: "
        self.lblGE.AutoSize = True
        self.gbG.Controls.Add(self.lblGE)

        self.lblGEv = Label()
        self.lblGEv.Left = self.lblGE.Left + self.lblGE.Width
        self.lblGEv.Top = 25
        self.lblGEv.Text = ""
        self.lblGEv.AutoSize = True
        self.gbG.Controls.Add(self.lblGEv)

        self.btnGE = Button()
        self.btnGE.Left = 10
        self.btnGE.Top = 40
        self.btnGE.Width = 140
        self.btnGE.Text = "Get GREEN Exposure"
        self.btnGE.Click += self.btnGE_Click
        self.gbG.Controls.Add(self.btnGE)

        self.lblGG = Label()
        self.lblGG.Left = 10
        self.lblGG.Top = 75
        self.lblGG.Text = "Gain: "
        self.lblGG.AutoSize = True
        self.gbG.Controls.Add(self.lblGG)

        self.lblGGv = Label()
        self.lblGGv.Left = self.lblGG.Left + self.lblGG.Width
        self.lblGGv.Top = 75
        self.lblGGv.Text = ""
        self.lblGGv.AutoSize = True
        self.gbG.Controls.Add(self.lblGGv)
        
        self.btnGG = Button()
        self.btnGG.Left = 10
        self.btnGG.Top = 90
        self.btnGG.Width = 140
        self.btnGG.Text = "Get GREEN Gain"
        self.btnGG.Click += self.btnGG_Click
        self.gbG.Controls.Add(self.btnGG)

        self.Controls.Add(self.gbG)

        # BLUE FILTER SECTION
        self.gbB = GroupBox()
        self.gbB.Left = 580
        self.gbB.Top = 25
        self.gbB.Width = 160
        self.gbB.Height = 130
        self.gbB.Text = "BLUE FILTER #4"

        self.lblBE = Label()
        self.lblBE.Left = 10
        self.lblBE.Top = 25
        self.lblBE.Text = "Exposure: "
        self.lblBE.AutoSize = True
        self.gbB.Controls.Add(self.lblBE)

        self.lblBEv = Label()
        self.lblBEv.Left = self.lblBE.Left + self.lblBE.Width
        self.lblBEv.Top = 25
        self.lblBEv.Text = ""
        self.lblBEv.AutoSize = True
        self.gbB.Controls.Add(self.lblBEv)

        self.btnBE = Button()
        self.btnBE.Left = 10
        self.btnBE.Top = 40
        self.btnBE.Width = 140
        self.btnBE.Text = "Get BLUE Exposure"
        self.btnBE.Click += self.btnBE_Click
        self.gbB.Controls.Add(self.btnBE)

        self.lblBG = Label()
        self.lblBG.Left = 10
        self.lblBG.Top = 75
        self.lblBG.Text = "Gain: "
        self.lblBG.AutoSize = True
        self.gbB.Controls.Add(self.lblBG)

        self.lblBGv = Label()
        self.lblBGv.Left = self.lblBG.Left + self.lblBG.Width
        self.lblBGv.Top = 75
        self.lblBGv.Text = ""
        self.lblBGv.AutoSize = True
        self.gbB.Controls.Add(self.lblBGv)
        
        self.btnBG = Button()
        self.btnBG.Left = 10
        self.btnBG.Top = 90
        self.btnBG.Width = 140
        self.btnBG.Text = "Get BLUE Gain"
        self.btnBG.Click += self.btnBG_Click
        self.gbB.Controls.Add(self.btnBG)

        self.Controls.Add(self.gbB)

        # TIME OR FRAME AND SEQUENCE SECTION
        self.gbTFS = GroupBox()
        self.gbTFS.Left = 25
        self.gbTFS.Top = 160
        self.gbTFS.Width = 715
        self.gbTFS.Height = 60
        self.gbTFS.Text = "TimeLimit or FrameLimit && Sequence Count"

        self.rbTFST = RadioButton()
        self.rbTFST.Left = 10
        self.rbTFST.Top = 20
        self.rbTFST.Text = "Time Limit"
        self.rbTFST.CheckedChanged += self.rbTF_Click
        self.gbTFS.Controls.Add(self.rbTFST)
        
        self.rbTFSF = RadioButton()
        self.rbTFSF.Left = 195
        self.rbTFSF.Top = 20
        self.rbTFSF.Text = "Frame Limit"
        self.rbTFSF.CheckedChanged += self.rbTF_Click
        self.gbTFS.Controls.Add(self.rbTFSF)

        self.lblTFSTF = Label()
        self.lblTFSTF.Left = 380
        self.lblTFSTF.Top = 15
        self.lblTFSTF.Width = 120
        self.lblTFSTF.Height = 13
        self.lblTFSTF.Text = "Time/Frame Limit"
        self.gbTFS.Controls.Add(self.lblTFSTF)

        self.nudTFSTF = NumericUpDown()
        self.nudTFSTF.Left = 408
        self.nudTFSTF.Top = 30
        self.nudTFSTF.Width = 44
        self.nudTFSTF.Value = 1
        self.nudTFSTF.Minimum = 1
        self.nudTFSTF.Maximum = 9999
        self.gbTFS.Controls.Add(self.nudTFSTF)

        self.lblTFSS = Label()
        self.lblTFSS.Left = 563
        self.lblTFSS.Top = 15
        self.lblTFSS.Width = 120
        self.lblTFSS.Height = 13
        self.lblTFSS.Text = "Sequences per Filter"
        self.gbTFS.Controls.Add(self.lblTFSS)

        self.nudTFSS = NumericUpDown()
        self.nudTFSS.Left = 593
        self.nudTFSS.Top = 30
        self.nudTFSS.Width = 44
        self.nudTFSS.Value = 1
        self.nudTFSS.Minimum = 1
        self.nudTFSS.Maximum = 9999
        self.gbTFS.Controls.Add(self.nudTFSS)

        self.Controls.Add(self.gbTFS)

        # PROGRESS SECTION
        self.gbProgress = GroupBox()
        self.gbProgress.Left = 25
        self.gbProgress.Top = 225
        self.gbProgress.Width = 715
        self.gbProgress.Height = 145
        self.gbProgress.Text = "Progress"

        self.tbOutput = TextBox()
        self.tbOutput.Left = 10
        self.tbOutput.Top = 25
        self.tbOutput.Width = 695
        self.tbOutput.Height = 100
        self.tbOutput.Multiline = True
        self.tbOutput.ReadOnly = True
        self.tbOutput.ScrollBars = ScrollBars.Vertical
        self.gbProgress.Controls.Add(self.tbOutput)

        self.Controls.Add(self.gbProgress)

        self.btnHelp = Button()
        self.btnHelp.Left = 25
        self.btnHelp.Top = 380
        self.btnHelp.Width = 85
        self.btnHelp.Text = "Help"
        self.btnHelp.Click += self.btnHelp_Click

        self.Controls.Add(self.btnHelp)

        self.btnSC = Button()
        self.btnSC.Left = 135
        self.btnSC.Top = 380
        self.btnSC.Width = 495
        self.btnSC.Text = "Start Capture"
        self.btnSC.Click += self.btnSC_Click

        self.Controls.Add(self.btnSC)

        self.btnAbout = Button()
        self.btnAbout.Left = 655
        self.btnAbout.Top = 380
        self.btnAbout.Width = 85
        self.btnAbout.Text = "About"
        self.btnAbout.Click += self.btnAbout_Click

        self.Controls.Add(self.btnAbout)

        # FUNCTIONS SECTION

    def btnHelp_Click(self, sender, event):
        self.appendMessageText('')
        self.appendMessageText('*** HELP ***')
        self.appendMessageText('')
        self.appendMessageText(' 1.FIRSTLY MAKE SURE THAT EACH FILTERS FOCUS OFFSET IS ADDED IN THE ASCOM FILTERWHEEL SETUP DIALOG UNLESS PARFOCAL FILTERS ARE BEING USED')
        self.appendMessageText('')
        self.appendMessageText(' 2.IT IS ALSO IMPORTANT THAT THE FILTERS ARE ARRANGED IN THE FILTERWHEEL AS L,R,G & B IN POSITIONS 1,2,3 & 4 RESPECTIVELY')
        self.appendMessageText('')
        self.appendMessageText(' 3.MAKE SURE THAT YOUR TARGET IS IN FOCUS & THAT A TARGET NAME IS PICKED OR ENTERED IN SHARPCAP')
        self.appendMessageText('')
        self.appendMessageText(' 4.START LRGB SEQUENCE CAPTURE FROM THE BUTTON ON THE TOOLBAR')
        self.appendMessageText('')
        self.appendMessageText(' 5.BEGIN BY MAKING SURE THAT FILTER [L] IS SELECTED IN YOUR FILTERWHEEL USING THE SHARPCAP FILTERWHEEL INTERFACE')
        self.appendMessageText('')
        self.appendMessageText(' 6.NOW USE THE EXPOSURE & GAIN SLIDERS IN SHARPCAP TO GET THE TARGET AT THE CORRECT SETTINGS FOR CAPTURE, CHECK HISTOGRAM IF REQUIRED')
        self.appendMessageText('')
        self.appendMessageText(' 7.ONCE THE REQUIRED EXPOSURE & GAIN ARE FOUND PRESS THE [Get LUM Exposure] & [Get LUM Gain] BUTTONS ON LRGB SEQUENCE CAPTURE')
        self.appendMessageText('')
        self.appendMessageText(' 8.REPEAT STEPS 5 TO 7 BY ITERATING THROUGH THE OTHER FILTERS IN THE FILTERWHEEL & PRESSING THE CORRESPONDING [Get Exposure] & [Get Gain] BUTTONS')
        self.appendMessageText('')
        self.appendMessageText(' 9.CHOOSE BETWEEN EITHER A TIMELIMITED OR FRAMELIMITED CAPTURE SEQUENCE BY SELECTING THE OPTION')
        self.appendMessageText('')
        self.appendMessageText('10.ENTER A VALUE FOR THE TIMELIMIT (in seconds) OR FRAMELIMIT (number of frames)')
        self.appendMessageText('')
        self.appendMessageText('11.ENTER A VALUE FOR THE NUMBER OF SEQUENCES TO CAPTURE')
        self.appendMessageText('')
        self.appendMessageText('12.HIT THE [START CAPTURE] BUTTON & WAIT FOR THE SEQUENCE TO FINISH')
        self.appendMessageText('')

    def btnAbout_Click(self, sender, event):
        self.appendMessageText('')
        self.appendMessageText('*** ABOUT ***')
        self.appendMessageText('')
        self.appendMessageText('JUST A MASSIVE THANK YOU TO')
        self.appendMessageText('')
        self.appendMessageText('ROBIN [admin] & BRIAN [oopfan]')
        self.appendMessageText('')
        self.appendMessageText('ON THE SHARPCAP FORUM FOR HELPING ME OUT WHEN I WAS STUCK')
        self.appendMessageText('')
        

    def btnLE_Click(self, sender, event):
        global expL
        expL = SharpCap.SelectedCamera.Controls.Exposure.Value
        self.lblLEv.Text = get_EV()
        SharpCap.ShowNotification('Got LUMINANCE Exposure Value of: ' + self.lblLEv.Text)
        self.appendMessageText('Got LUMINANCE Exposure Value of: ' + self.lblLEv.Text)
        

    def btnLG_Click(self, sender, event):
        GV = SharpCap.SelectedCamera.Controls.Gain.Value
        self.lblLGv.Text = str(GV)
        global gainL
        gainL = GV
        SharpCap.ShowNotification('Got LUMINANCE Gain Value of: ' + self.lblLGv.Text)
        self.appendMessageText('Got LUMINANCE Gain Value of: ' + self.lblLGv.Text)
        

    def btnRE_Click(self, sender, event):
        global expR
        expR = SharpCap.SelectedCamera.Controls.Exposure.Value
        self.lblREv.Text = get_EV()
        SharpCap.ShowNotification('Got RED Exposure Value of: ' + self.lblREv.Text)
        self.appendMessageText('Got RED Exposure Value of: ' + self.lblREv.Text)
        

    def btnRG_Click(self, sender, event):
        GV = SharpCap.SelectedCamera.Controls.Gain.Value
        self.lblRGv.Text = str(GV)
        global gainR
        gainR = GV
        SharpCap.ShowNotification('Got RED Gain Value of: ' + self.lblRGv.Text)
        self.appendMessageText('Got RED Gain Value of: ' + self.lblRGv.Text)
        

    def btnGE_Click(self, sender, event):
        global expG
        expG = SharpCap.SelectedCamera.Controls.Exposure.Value
        self.lblGEv.Text = get_EV()
        SharpCap.ShowNotification('Got GREEN Exposure Value of: ' + self.lblGEv.Text)
        self.appendMessageText('Got GREEN Exposure Value of: ' + self.lblGEv.Text)
        

    def btnGG_Click(self, sender, event):
        GV = SharpCap.SelectedCamera.Controls.Gain.Value
        self.lblGGv.Text = str(GV)
        global gainG
        gainG = GV
        SharpCap.ShowNotification('Got GREEN Gain Value of: ' + self.lblGGv.Text)
        self.appendMessageText('Got GREEN Gain Value of: ' + self.lblGGv.Text)
        

    def btnBE_Click(self, sender, event):
        global expB
        expB = SharpCap.SelectedCamera.Controls.Exposure.Value
        self.lblBEv.Text = get_EV()
        SharpCap.ShowNotification('Got BLUE Exposure Value of: ' + self.lblBEv.Text)
        self.appendMessageText('Got BLUE Exposure Value of: ' + self.lblBEv.Text)
        

    def btnBG_Click(self, sender, event):
        GV = SharpCap.SelectedCamera.Controls.Gain.Value
        self.lblBGv.Text = str(GV)
        global gainB
        gainB = GV
        SharpCap.ShowNotification('Got BLUE Gain Value of: ' + self.lblBGv.Text)
        self.appendMessageText('Got BLUE Gain Value of: ' + self.lblBGv.Text)
        

    def rbTF_Click(self, sender, event):
        global capTYPE
        if sender.Checked:
            SharpCap.ShowNotification('Your Chosen Option Is: ' + sender.Text)
            self.appendMessageText('Your Chosen Option Is: ' + sender.Text)
            if sender.Text != 'Frame Limit':
                self.lblTFSTF.Text = 'Time Limit per Filter (s)'
                self.lblTFSTF.Left = 375
                capTYPE = 0
            else:
                self.lblTFSTF.Text = 'Frame Limit per Filter'
                self.lblTFSTF.Left = 378
                capTYPE = 1

    def btnSC_Click(self, sender, event):
        if not any([self.rbTFST.Checked, self.rbTFSF.Checked]):
            SharpCap.ShowNotification('You Need To Choose Either The TimeLimited or FrameLimited Option.')
            MessageBox.Show('You Need To Choose Either The TimeLimited\n\nor FrameLimited Option!', 'ACTION REQUIRED', MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        else:
            self.chk_ExpEntries()

    def chk_ExpEntries(self):
        if not all([self.lblLEv.Text, self.lblREv.Text, self.lblGEv.Text, self.lblBEv.Text]):
            SharpCap.ShowNotification('There Are 1 Or More Missing Entries For The Exposure Values, Please Check.')
            MessageBox.Show('There Are 1 Or More Missing Entries For\n\nThe Exposure Values, Please Check!', 'ACTION REQUIRED', MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        else:
            self.chk_GainEntries()

    def chk_GainEntries(self):
        if not all([self.lblLGv.Text, self.lblRGv.Text, self.lblGGv.Text, self.lblBGv.Text]):
            SharpCap.ShowNotification('There Are 1 Or More Missing Entries For The Gain Values, Please Check.')
            MessageBox.Show('There Are 1 Or More Missing Entries For\n\nThe Gain Values, Please Check!', 'ACTION REQUIRED', MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        else:
            self.get_Values()

    def get_Values(self):
        global capTF
        capTF = int(float(self.nudTFSTF.Value))

        global capSeq
        capSeq = int(float(self.nudTFSS.Value))
        
        self.get_Target()
        
    def disable_GUI(self):
        self.ControlBox = False
        self.gbL.Enabled = False
        self.gbR.Enabled = False
        self.gbG.Enabled = False
        self.gbB.Enabled = False
        self.gbTFS.Enabled = False
        self.btnSC.Enabled = False
        self.btnAbout.Enabled = False
        self.btnHelp.Enabled = False
        LRGB_Custom_Button.Enabled = False

    def enable_GUI(self):
        self.ControlBox = True
        self.gbL.Enabled = True
        self.gbR.Enabled = True
        self.gbG.Enabled = True
        self.gbB.Enabled = True
        self.gbTFS.Enabled = True
        self.btnSC.Enabled = True
        self.btnAbout.Enabled = True
        self.btnHelp.Enabled = True
        LRGB_Custom_Button.Enabled = True

    def get_Target(self):
        tName = SharpCap.TargetName
        if tName is None:
            SharpCap.ShowNotification('NO TARGET NAME SET!')
            self.appendMessageText('NO TARGET NAME SET!')
            MessageBox.Show('NO TARGET NAME SET!', 'ACTION REQUIRED', MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            self.Text = 'LRGB Sequence Capture - NO TARGET NAME SET!'
            return
        else:
            SharpCap.ShowNotification('Target Name Set To: ' + tName)
            self.appendMessageText('Target Name Set To: ' + tName)
            self.Text = 'LRGB Sequence Capture - Target: ' + tName
            if capTYPE != 1:
                worker = tWorkerThread(self)
                delegate = ThreadStart(worker.CapDO)
                thread = Thread(delegate)
                daemon = True
                thread.Start()
            else:
                worker = fWorkerThread(self)
                delegate = ThreadStart(worker.CapDO)
                thread = Thread(delegate)
                daemon = True
                thread.Start()
            
    def setup_tCapture(self, Exp, Gain, Seconds):
        SharpCap.SelectedCamera.Controls.Exposure.Value = Exp
        SharpCap.SelectedCamera.Controls.Gain.Value = Gain
        SharpCap.SelectedCamera.CaptureConfig.CaptureLimitType = CaptureLimitType.TimeLimited
        if SCVer < 3.2:
            SharpCap.SelectedCamera.CaptureConfig.CaptureLimitValue = Seconds
        else:
            SharpCap.SelectedCamera.CaptureConfig.CaptureLimitTime = TimeSpan.FromSeconds(Seconds)
        time.sleep(2)
        SharpCap.SelectedCamera.PrepareToCapture()
        SharpCap.SelectedCamera.RunCapture()
        while True:
            if not SharpCap.SelectedCamera.Capturing :
                break
            time.sleep(0.5)

    def get_tCapture(self):
        self.disable_GUI()
        SharpCap.ShowNotification('Moving Filter Wheel To LUMINANCE Position To Start Capture Sequence.')
        self.appendMessageText('Moving Filter Wheel To LUMINANCE Position To Start Capture Sequence.')
        SharpCap.Wheels.SelectedWheel.Position = 1
        time.sleep(1)

        for count in range(1,capSeq+1):
            SharpCap.ShowNotification('Capturing TimeLimited LUMINANCE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing TimeLimited LUMINANCE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 1
            self.setup_tCapture(expL, gainL, capTF)
            self.appendMessageText('Moving Filter Wheel To RED Position.')
            SharpCap.ShowNotification('Capturing TimeLimited RED Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing TimeLimited RED Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 2
            self.setup_tCapture(expR, gainR, capTF)
            self.appendMessageText('Moving Filter Wheel To GREEN Position.')
            SharpCap.ShowNotification('Capturing TimeLimited GREEN Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing TimeLimited GREEN Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 3
            self.setup_tCapture(expG, gainG, capTF)
            self.appendMessageText('Moving Filter Wheel To BLUE Position.')
            SharpCap.ShowNotification('Capturing TimeLimited BLUE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing TimeLimited BLUE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 4
            self.setup_tCapture(expB, gainB, capTF)
            if count != capSeq:
                self.appendMessageText('Moving Filter Wheel To LUMINANCE Position')
            else:
                SharpCap.ShowNotification('Capture Sequence Completed.')
                self.appendMessageText('Capture Sequence Completed.')
                self.appendMessageText('Moving Filter Wheel To LUMINANCE Parking Position.')
                SharpCap.Wheels.SelectedWheel.Position = 1
                break
        self.enable_GUI()

    def setup_fCapture(self, Exp, Gain, Frames):
        SharpCap.SelectedCamera.Controls.Exposure.Value = Exp
        SharpCap.SelectedCamera.Controls.Gain.Value = Gain
        SharpCap.SelectedCamera.CaptureConfig.CaptureLimitType = CaptureLimitType.FrameLimited
        if SCVer < 3.2:
            SharpCap.SelectedCamera.CaptureConfig.CaptureLimitValue = Frames
        else:
            SharpCap.SelectedCamera.CaptureConfig.CaptureLimitCount = Frames
        time.sleep(2)
        SharpCap.SelectedCamera.PrepareToCapture()
        SharpCap.SelectedCamera.RunCapture()
        while True:
            if not SharpCap.SelectedCamera.Capturing :
                break
            time.sleep(0.5)

    def get_fCapture(self):
        self.disable_GUI()
        SharpCap.ShowNotification('Moving Filter Wheel To LUMINANCE Position To Start Capture Sequence.')
        self.appendMessageText('Moving Filter Wheel To LUMINANCE Position To Start Capture Sequence.')
        SharpCap.Wheels.SelectedWheel.Position = 1
        time.sleep(1)

        for count in range(1,capSeq+1):
            SharpCap.ShowNotification('Capturing FrameLimited LUMINANCE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing FrameLimited LUMINANCE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 1
            self.setup_fCapture(expL, gainL, capTF)
            self.appendMessageText('Moving Filter Wheel To RED Position.')
            SharpCap.ShowNotification('Capturing FrameLimited RED Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing FrameLimited RED Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 2
            self.setup_fCapture(expR, gainR, capTF)
            self.appendMessageText('Moving Filter Wheel To GREEN Position.')
            SharpCap.ShowNotification('Capturing FrameLimited GREEN Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing FrameLimited GREEN Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 3
            self.setup_fCapture(expG, gainG, capTF)
            self.appendMessageText('Moving Filter Wheel To BLUE Position.')
            SharpCap.ShowNotification('Capturing FrameLimited BLUE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            self.appendMessageText('Capturing FrameLimited BLUE Channel: (' + str(count) + '/' + str(capSeq) + ')')
            SharpCap.Wheels.SelectedWheel.Position = 4
            self.setup_fCapture(expB, gainB, capTF)
            if count != capSeq:
                self.appendMessageText('Moving Filter Wheel To LUMINANCE Position')
            else:
                SharpCap.ShowNotification('Capture Sequence Completed.')
                self.appendMessageText('Capture Sequence Completed.')
                self.appendMessageText('Moving Filter Wheel To LUMINANCE Parking Position.')
                SharpCap.Wheels.SelectedWheel.Position = 1
                break
        self.enable_GUI()
                
    def appendMessageText(self, text):
        self.tbOutput.AppendText(text + Environment.NewLine)

def launch_form():
    if all([SharpCap.Wheels.SelectedWheel.Connected, SharpCap.SelectedCamera]):
        form = Main()
        form.StartPosition = FormStartPosition.CenterScreen
        form.Show()
    else:
        SharpCap.ShowNotification('ONE or BOTH DEVICES [Camera or FilterWheel] NOT CONNECTED')
        MessageBox.Show('ONE or BOTH DEVICES [Camera or FilterWheel] NOT CONNECTED!', 'ACTION REQUIRED', MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

    

def get_EV():
    EV =  SharpCap.SelectedCamera.Controls.Exposure.Value
    if EV >= 50:
        return str(format(EV, '.0f')) + " s"
    elif EV >= 1.5 and EV < 50:
        return str(format(EV, '.1f')) + " s"
    elif EV >= 0.05 and EV < 1.5:
        return str(format(EV * 1000, '.0f')) + " ms"
    elif EV >= 0.001 and EV < 0.05:
        return str(format(EV * 1000, '.1f')) + " ms"
    elif EV >= 0.0001 and EV < 0.001:
        return str(format(EV * 1000, '.2f')) + " ms"
    else:
        return str(format(EV * 1000, '.3f')) + " ms"

class tWorkerThread:
    def __init__(self, form):
        self.form = form

    def CapDO(self):
        self.form.get_tCapture()

class fWorkerThread:
    def __init__(self, form):
        self.form = form

    def CapDO(self):
        self.form.get_fCapture()

LRGB_Custom_Button = SharpCap.AddCustomButton("LRGB Sequence Capture", None, "LRGB Sequence Capture", launch_form)
