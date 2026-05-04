const express = require('express');
const fs = require('fs');
const { Packer } = require('docx');
const path = require('path');

// This imports the function from your original script
const { createDoc } = require('./generateGuide');

const app = express();
const PORT = process.env.PORT || 10000; // Use the port assigned by Render

app.get('/generate-guide', async (req, res) => {
    try {
        // Call the function from your script to get the document buffer
        const docBuffer = await createDoc();

        // Set the correct headers for a file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="EDI_BA_JobReady_Guide_2026.docx"');
        res.send(docBuffer);
        console.log('Guide generated and download started successfully.');
    } catch (error) {
        console.error('Error generating guide:', error);
        res.status(500).send('Failed to generate the document.');
    }
});

// A simple route to show the service is running
app.get('/', (req, res) => {
    res.send(`
        <h1>EDI Guide Generator is Running!</h1>
        <p>Click the link below to generate the guide:</p>
        <a href="/generate-guide">Download EDI BA Job-Ready Guide</a>
    `);
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});
