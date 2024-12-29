import { NextRequest, NextResponse } from 'next/server';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import * as fs from 'fs';
import * as path from 'path';

import mammoth from 'mammoth';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
    try {
        const data = await request.json();
        const format = request.nextUrl.searchParams.get('format') || 'pdf';

        console.log('Received data:', data);

        const templatePath = path.join(process.cwd(), 'public', 'word', 'Lulab_invioce.docx');
        if (!fs.existsSync(templatePath)) {
            return NextResponse.json({ error: 'Template file not found' }, { status: 404 });
        }

        const template = fs.readFileSync(templatePath);
        const zip = new PizZip(template);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        try {
            doc.render(data);
        } catch (error) {
            console.error("Error during template rendering:", error);
            return NextResponse.json({ error: 'Template rendering failed' }, { status: 500 });
        }

        const wordContent = doc.getZip().generate({
            type: 'nodebuffer',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        if (format === 'word') {
            return new NextResponse(wordContent, {
                headers: {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'Content-Disposition': 'attachment; filename="Generated_Document.docx"',
                },
            });
        }

        // Use '/tmp' directory for temporary file storage in Vercel
        const tempDir = '/tmp'; // Update to /tmp
        // Vercel already has this directory, so no need to check for existence
        // Ensure temp directory exists
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir, { recursive: true });
        }

        // Save generated DOCX to a temporary file
        const tempDocxPath = path.join(tempDir, `temp_${Date.now()}.docx`);
        fs.writeFileSync(tempDocxPath, wordContent);

        // Convert DOCX to HTML using mammoth.js
        let htmlContent;
        try {
            htmlContent = await convertDocxToHtml(tempDocxPath);
        } catch (error) {
            console.error('Error converting DOCX to HTML:', error);
            fs.unlinkSync(tempDocxPath); // Clean up temp DOCX file
            return NextResponse.json({ error: 'Conversion from DOCX to HTML failed' }, { status: 500 });
        }

        // Convert HTML to PDF using Puppeteer
        let pdfBuffer;
        try {
            pdfBuffer = await convertHtmlToPdf(htmlContent);
        } catch (error) {
            console.error('Error converting HTML to PDF:', error);
            fs.unlinkSync(tempDocxPath); // Clean up temp DOCX file
            return NextResponse.json({ error: 'Conversion from HTML to PDF failed' }, { status: 500 });
        }

        // Clean up temporary DOCX file
        fs.unlinkSync(tempDocxPath);

        // Return PDF file as response (convert pdfBuffer to Buffer)
        return new NextResponse(Buffer.from(pdfBuffer), {
            headers: {
                'Content-Type': 'application/pdf',
                'Content-Disposition': 'attachment; filename="Generated_Document.pdf"',
            },
        });

    } catch (error) {
        console.error("General error:", error);
        return NextResponse.json({ error: (error as Error).message }, { status: 500 });
    }
}


// Convert DOCX to HTML using mammoth.js
async function convertDocxToHtml(docxPath: string): Promise<string> {
    const docxBuffer = fs.readFileSync(docxPath);
    const { value: html } = await mammoth.convertToHtml({ buffer: docxBuffer });
    return html;
}

// Convert HTML to PDF using Puppeteer
function convertHtmlToPdf(html: string): Promise<Buffer> {
    return new Promise(async (resolve, reject) => {
        try {
            const browser = await puppeteer.launch({
                headless: true, // In production, should run in headless mode
                args: ['--no-sandbox', '--disable-setuid-sandbox'], // Useful for Docker or restricted environments
            });
            const page = await browser.newPage();
            await page.setContent(html);
            const pdfBuffer = await page.pdf({ format: 'A4' });
            await browser.close();
            resolve(Buffer.from(pdfBuffer)); // Ensure pdfBuffer is returned as Buffer
        } catch (error) {
            reject(error);
        }
    });
}
