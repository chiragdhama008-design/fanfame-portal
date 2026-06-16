const express = require('express');
const nodemailer = require('nodemailer');
const DocxTemplateModule = require('docx-templates');
const DocxTemplate = DocxTemplateModule.DocxTemplate || DocxTemplateModule.default || DocxTemplateModule;
const fs = require('fs');
const path = require('path');
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
    const templatePath = path.join(process.cwd(), 'OfferLetter.docx');
    
    if (!fs.existsSync(templatePath)) {
        return res.status(500).json({ error: 'OfferLetter.docx not found in server root.' });
    }
    
    const templateBuffer = fs.readFileSync(templatePath);

    const updatedDocBuffer = await DocxTemplate.create({
      template: templateBuffer,
      data: {
        employee_name: name,
        start_date: startDate,
        salary_amount: salary,
      },
    });

    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS, 
      },
    });

    const emailHtml = `
      <p>Hi ${name},</p>
      <p>We're happy to offer you the position of <strong>Twitter Growth Assistant</strong> at FanFame Media.</p>
      <p>Your salary will be <strong>₹${salary}</strong> per month.</p>
      <p>As a Twitter Growth Assistant, you will be responsible for managing assigned Twitter/X accounts and following lead lists provided by the company. Training and clear instructions will be provided, and your role is to complete the assigned daily tasks consistently and accurately.</p>
      <p>We're looking for someone who is reliable, consistent, able to follow instructions, comfortable with repetitive tasks, and communicates professionally.</p>
      <p>Any company accounts, systems, SOPs, lead lists, or internal information shared with you must remain confidential and may not be shared with anyone outside the company.</p>
      <p>To accept this offer, please review the attached contract, sign it, and return the signed copy by replying to this email.</p>
      <p>Upon receipt of the signed contract, we will provide further onboarding instructions and confirm your start date.</p>
      <p>We're excited to have you join the team.</p>
      <br>
      <p>Best regards,</p>
      <p><strong>George</strong><br>FanFame Media</p>
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
    return res.status(500).json({ error: error.message || 'Server error generating document.' });
  }
});

module.exports = app;
