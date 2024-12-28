import { NextRequest, NextResponse } from 'next/server';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';

export async function POST(request: NextRequest) {
    try {
        const data = await request.json();
        const format = request.nextUrl.searchParams.get('format') || 'pdf';

        // Check if the data contains any unexpected or special characters
        console.log('Received data:', data);

        // Read template file (Word template part remains unchanged)
        const templatePath = path.join(process.cwd(), 'public', 'word', 'Lulab_invioce.docx');
        if (!fs.existsSync(templatePath)) {
            return NextResponse.json({ error: 'Template file not found' }, { status: 404 });
        }

        const template = fs.readFileSync(templatePath);
        const zip = new PizZip(template);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Fill template data (Word part remains unchanged)
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

        // If the user wants the Word document (Word part remains unchanged)
        if (format === 'word') {
            return new NextResponse(wordContent, {
                headers: {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'Content-Disposition': 'attachment; filename="Generated_Document.docx"',
                },
            });
        }

        // Ensure temp directory exists
        const tempDir = path.join(process.cwd(), 'temp');
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir, { recursive: true });  // Ensure directory exists
        }

        // Convert to PDF using LibreOffice's soffice
        const tempDocxPath = path.join(tempDir, `temp_${Date.now()}.docx`);
        fs.writeFileSync(tempDocxPath, wordContent);

        const pdfFilePath = path.join(tempDir, `temp_${Date.now()}.pdf`); // Use same name as the generated PDF file

        try {
            // Updated path for soffice.wrapper.sh
            const sofficePath = '/opt/homebrew/bin/soffice';  // 使用正确的 soffice 路径

            // Use soffice to convert DOCX to PDF
            await new Promise<void>((resolve, reject) => {
                exec(`${sofficePath} --headless --convert-to pdf --outdir ${tempDir} ${tempDocxPath}`, (error, stdout, stderr) => {
                    if (error) {
                        console.error(`Error during PDF conversion: ${stderr}`);
                        reject(stderr);
                    } else {
                        console.log(`PDF conversion succeeded: ${stdout}`);
                        resolve();
                    }
                });
            });

            // Check if the PDF file exists before reading
            if (fs.existsSync(pdfFilePath)) {
                // Read the generated PDF
                const pdfBuffer = fs.readFileSync(pdfFilePath);

                // Clean up temp files
                fs.unlinkSync(tempDocxPath);
                fs.unlinkSync(pdfFilePath);

                return new NextResponse(pdfBuffer, {
                    headers: {
                        'Content-Type': 'application/pdf',
                        'Content-Disposition': 'attachment; filename="Generated_Document.pdf"',
                    },
                });
            } else {
                console.error('PDF file was not generated:', pdfFilePath);
                return NextResponse.json({ error: 'PDF conversion failed - file not found' }, { status: 500 });
            }
        } catch (err) {
            // Clean up temp file in case of error
            fs.unlinkSync(tempDocxPath);
            console.error("Error during PDF conversion:", err);
            return NextResponse.json({ error: 'PDF conversion failed' }, { status: 500 });
        }
    } catch (error) {
        console.error("General error:", error);
        return NextResponse.json({ error: (error as Error).message }, { status: 500 });
    }
}
