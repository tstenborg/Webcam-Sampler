# Webcam-Sampler

Automated sampling of webcam imaging via Excel VBA.

---

<figure style="width:416px;">
  <img src="WebcamCollage.png" alt="Sampling webcams managed by Airservices Australia on Norfolk Island." width="416" height="360">
  <figcaption>Figure 1. Sampling webcams managed by Airservices Australia on Norfolk Island.</figcaption>
</figure>

---

### Key Files

- Automated Webcam Sampler.xlsm &nbsp;&nbsp; Macro-enabled Excel workbook.<br />
- WebcamSampler.bas &nbsp;&nbsp; VBA module.<br />

Automated sampling of webcam imaging was investigated as a means of meteor data collection. A sampling system was built with Visual Basic for Applications (VBA), the scripting language built into Microsoft's Office suite.

The Excel workbook runs a standalone automated webcam sampler. Its constituent VBA module, WebcamSampler.bas, has also been provided as separate file for a) easy code review outside Excel and b) importing into another Excel workbook, if desired.

### Software Requirements

- Excel (Windows version).

The webcam sampler has been tested with Excel, version 2508, on Windows 11.

### Sampler Configuration

The webcam sampler has default settings for target webcams, sampling cadence, sampling period, download directory, etc. These should be configured in VBA to meet your needs prior to executing the sampler. A modest, not advanced, level of VBA knowledge is needed for this configuration.

### Execution Options

Ensure VBA macro execution is enabled in Excel. Once complete, there are three main ways to execute the sampler.

#### 1. Macro Button

The single Excel worksheet features a "Take Shots" button at the top left of the worksheet that will execute the sampler.

#### 2. VBA Editor

The DownloadFileAPI function of the VBA code will execute webcam sampling. This can be triggered from Excel's VBA editor. The VBA editor can be opened by selecting Developer, Visual Basic from the Excel ribbon. If the Developer option isn't visible in the Excel ribbon, it can be added by selecting File, Options, Customize Ribbon from the Excel main menu. (Excel version 2508 menu options.)

#### 3. Imported Module

To execute the webcam sampler from another Excel workbook, a VBA module, WebcamSampler.bas, is available for import and execution. The sampler configuration options will still need to be customised to meet your sampling needs.

### Reference

Stenborg, TN 2019, "[Meteor Candidate Observations from Automated Weather Camera Sampling in VBA](https://www.saasst.ae/images/spsimpleportfolio/uaemmn/IMC2019-Proceedings2019.pdf#page=193)", in U Pajer, J Rendtel, M Gyssens and C Verbeeck (eds), Proceedings of the International Meteor Conference, Bollmannsruh, Germany, pp. 189&ndash;190.
