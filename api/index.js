const express = require('express');
const nodemailer = require('nodemailer');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = require('docx');
const cors = require('cors');

const app = express();
app.use(express.json());
app.use(cors());

app.post('/api/send-offer', async (req, res) => {
  const { name, email, startDate, salary } = req.body;

  if (!name || !email || !startDate || !salary) {
    return res.status(400).json({ error: 'All fields are required.' });
  }

  try {
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: "TWITTER GROWTH ASSISTANT AGREEMENT",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: "" }),
          new Paragraph({
            children: [
              new TextRun({ text: `This Agreement is effective as of ${startDate} and is entered into between FanFame Media and ${name}.`, bold: false }),
            ],
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "Both parties agree to the following terms and conditions:", heading: HeadingLevel.NORMAL }),
          new Paragraph({ text: "" }),
          
          new Paragraph({ text: "1. Position", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The Company hereby engages the Assistant as a Twitter Growth Assistant. The Assistant will perform services on behalf of the Company and agrees to carry out all assigned duties in a professional and responsible manner." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "2. Duties and Responsibilities", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The Assistant will be responsible for managing Twitter/X accounts assigned by the Company. Responsibilities may include:" }),
          new Paragraph({ text: "• Following lead lists provided by the Company" }),
          new Paragraph({ text: "• Completing assigned daily tasks and quotas" }),
          new Paragraph({ text: "• Following Company SOPs and account safety guidelines" }),
          new Paragraph({ text: "• Reporting account restrictions, issues, or unusual activity" }),
          new Paragraph({ text: "• Participating in training and team meetings when required" }),
          new Paragraph({ text: "The Assistant agrees to follow all instructions provided by management and to perform duties accurately and consistently." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "3. Work Schedule", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The Assistant agrees to work six (6) days per week. Working hours and schedules will be communicated by management and may be adjusted as business needs require. The Assistant is expected to maintain regular communication and remain available during scheduled working hours." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "4. Probationary Period", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The first seven (7) days of employment shall be considered a probationary period. During this period, either party may terminate this Agreement at any time without notice. Successful completion of the probationary period does not guarantee continued employment but confirms that the Assistant has met the Company's initial performance expectations." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "5. Compensation", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({
            children: [
              new TextRun({ text: `The Assistant shall receive a fixed salary of ₹${salary} per month. Any future salary increases, bonuses, or additional compensation shall be determined solely at the discretion of the Company based on performance, reliability, attendance, and overall contribution.`, bold: false }),
            ],
          }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "6. Training, Equipment and Internet Requirements", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The Assistant agrees to participate in training sessions, meetings, screen-sharing sessions, or other coaching activities when requested by management. The Assistant is responsible for maintaining a reliable computer or laptop, a stable internet connection, and any equipment necessary to perform assigned duties." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "7. Confidentiality", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "The Assistant acknowledges that they may receive access to confidential Company information, including login credentials, SOPs, lead lists, client information, and internal communications. The Assistant agrees not to disclose, share, copy, distribute, or misuse any confidential information during or after employment." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "8. Termination", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "Following the probationary period, either party may terminate this Agreement by providing seven (7) days written notice. The Company reserves the right to terminate this Agreement immediately in cases of misconduct, repeated failure to perform duties, or breach of confidentiality." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "9. Entire Agreement", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "This Agreement represents the entire understanding between the parties and supersedes all previous discussions, communications, or agreements." }),
          new Paragraph({ text: "" }),

          new Paragraph({ text: "10. Acceptance", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: "By signing below, both parties acknowledge that they have read, understood, and agreed to the terms of this Agreement." }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: `Assistant Name: ${name}` }),
          new Paragraph({ text: `Date: ${startDate}` }),
          new Paragraph({ text: "Signature: __________________" }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "George (FanFame Media)" }),
          new Paragraph({ text: `Date: ${startDate}` }),
          new Paragraph({ text: "Signature: George" }),
        ],
      }],
    });

    const updatedDocBuffer = await Packer.toBuffer(doc);

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS, 
      },
    });

    // Clean HTML without bold tags
    const emailHtml = `
      <p>Hi ${name},</p>
      <p>We're happy to offer you the position of Twitter Growth Assistant at FanFame Media.</p>
      <p>Your salary will be ₹${salary} per month.</p>
      <p>As a Twitter Growth Assistant, you will be responsible for managing assigned Twitter/X accounts and following lead lists provided by the company. Training and clear instructions will be provided, and your role is to complete the assigned daily tasks consistently and accurately.</p>
      <p>We're looking for someone who is reliable, consistent, able to follow instructions, comfortable with repetitive tasks, and communicates professionally.</p>
      <p>Any company accounts, systems, SOPs, lead lists, or internal information shared with you must remain confidential and may not be shared with anyone outside the company.</p>
      <p>To accept this offer, please review the attached contract, sign it, and return the signed copy by replying to this email.</p>
      <p>Upon receipt of the signed contract, we will provide further onboarding instructions and confirm your start date.</p>
      <p>We're excited to have you join the team.</p>
      <br>
      <p>Best regards,</p>
      <p>George<br>FanFame Media</p>
    `;

    const mailOptions = {
      from: `"George - FanFame Media" <${process.env.EMAIL_USER}>`,
      to: email,
      subject: 'Offer letter - twitter management, FanFame media',
      html: emailHtml,
      attachments: [
        {
          filename: `OfferLetter_${name.replace(/\s+/g, '_')}.docx`,
          content: updatedDocBuffer,
        },
      ],
    };

    await transporter.sendMail(mailOptions);
    return res.status(200).json({ message: 'Success' });

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: 'Server error generating document.' });
  }
});

module.exports = app;
