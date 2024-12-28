'use client';

import { useState } from "react";

interface FormDataType {
    issuanceDate: string;
    address: string;
    city: string;
    state: string;
    postalCode: string;
    country: string;
    name: string;
    studentID: string;
    programName: string;
    startDate: string;
    endDate: string;
    tuitionFeeUSD: string;
}

const countries = [
    "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", 
    "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", 
    "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", 
    "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", 
    "Canada", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Congo-Brazzaville)", 
    "Costa Rica", "Croatia", "Cuba", "Cyprus", "Czechia (Czech Republic)", "Denmark", "Djibouti", 
    "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", 
    "Estonia", "Eswatini (fmr. 'Swaziland')", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", 
    "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", 
    "Haiti", "Holy See", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", 
    "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait", "Kyrgyzstan", 
    "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", 
    "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", 
    "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", 
    "Mozambique", "Myanmar (formerly Burma)", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", 
    "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway", "Oman", "Pakistan", 
    "Palau", "Palestine State", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", 
    "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", 
    "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", 
    "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", 
    "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", 
    "Sweden", "Switzerland", "Syria", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo", 
    "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", 
    "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", 
    "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"
];

export default function HomeContent() {
    const [formData, setFormData] = useState<FormDataType>({
        name: "",
        country: "China",
        state: "",
        city: "",
        postalCode: "",
        address: "",
        studentID: "",
        programName: "Practical Training Club",
        issuanceDate: "",
        startDate: "",
        endDate: "",
        tuitionFeeUSD: "",
    });
    const [generating, setGenerating] = useState(false);
    const [currentFormat, setCurrentFormat] = useState<"pdf" | "word" | null>(null);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
        const { name, value } = e.target;
        setFormData((prev) => ({ ...prev, [name]: value }));
    };

    const generateDocument = async (format: "pdf" | "word") => {
        if (generating) return; // Prevent further clicks while generating
        setGenerating(true);
        setCurrentFormat(format);
    
        try {
            const response = await fetch(`/api/generate_document?format=${format}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(formData),
            });
    
            if (!response.ok) throw new Error(await response.text());
    
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const fileExtension = format === "pdf" ? "pdf" : "docx"; // Set extension based on selected format
            const a = document.createElement("a");
            a.href = url;
            a.download = `Generated_Document.${fileExtension}`; // Dynamic file extension
            a.click();
        } catch (err) {
            alert((err as Error).message);
        } finally {
            setGenerating(false);
        }
    };
    
    return (
        <main style={{ padding: "2rem", maxWidth: "700px", margin: "auto" }}>
            <h1 style={{ textAlign: "center", fontSize: "2rem", marginBottom: "1.5rem" }}>Generate Offer Letter</h1>
            <form
                style={{
                    background: "#fff",
                    padding: "1.5rem",
                    borderRadius: "8px",
                    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.1)",
                }}
            >
                <div style={{ display: "grid", gap: "1.5rem" }}>
                    <div>
                        <label style={styles.label}>Name</label>
                        <input
                            type="text"
                            name="name"
                            value={formData.name}
                            onChange={handleChange}
                            placeholder="Enter full name"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Country</label>
                        <select
                            name="country"
                            value={formData.country}
                            onChange={handleChange}
                            style={styles.select}
                        >
                            {countries.map((c) => (
                                <option key={c} value={c}>
                                    {c}
                                </option>
                            ))}
                        </select>
                    </div>
                    <div>
                        <label style={styles.label}>State</label>
                        <input
                            type="text"
                            name="state"
                            value={formData.state}
                            onChange={handleChange}
                            placeholder="Enter state"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>City</label>
                        <input
                            type="text"
                            name="city"
                            value={formData.city}
                            onChange={handleChange}
                            placeholder="Enter city"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Address</label>
                        <input
                            type="text"
                            name="address"
                            value={formData.address}
                            onChange={handleChange}
                            placeholder="Enter address"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Postal Code</label>
                        <input
                            type="text"
                            name="postalCode"
                            value={formData.postalCode}
                            onChange={handleChange}
                            placeholder="Enter postal code"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Issuance Date</label>
                        <input
                            type="date"
                            name="issuanceDate"
                            value={formData.issuanceDate}
                            onChange={handleChange}
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Student ID</label>
                        <input
                            type="text"
                            name="studentID"
                            value={formData.studentID}
                            onChange={handleChange}
                            placeholder="Enter student ID"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Program Name</label>
                        <input
                            type="text"
                            name="programName"
                            value={formData.programName}
                            onChange={handleChange}
                            placeholder="Enter program name"
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Start Date</label>
                        <input
                            type="date"
                            name="startDate"
                            value={formData.startDate}
                            onChange={handleChange}
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>End Date</label>
                        <input
                            type="date"
                            name="endDate"
                            value={formData.endDate}
                            onChange={handleChange}
                            style={styles.input}
                        />
                    </div>
                    <div>
                        <label style={styles.label}>Tuition Fee (USD)</label>
                        <input
                            type="number"
                            name="tuitionFeeUSD"
                            value={formData.tuitionFeeUSD}
                            onChange={handleChange}
                            placeholder="Enter tuition fee"
                            style={styles.input}
                        />
                    </div>

                    <div style={{ display: "flex", justifyContent: "space-between", gap: "1rem" }}>
                    <button
                        type="button"
                        onClick={() => generateDocument("pdf")}
                        disabled={generating}
                        style={{ ...styles.button, backgroundColor: currentFormat === "pdf" ? "#0056b3" : "#007bff" }}
                    >
                        {generating && currentFormat === "pdf" ? "Generating PDF..." : "Generate PDF"}
                    </button>

                        <button
                            type="button"
                            onClick={() => generateDocument("word")}
                            disabled={generating}
                            style={{ ...styles.button, backgroundColor: currentFormat === "word" ? "#0056b3" : "#007bff" }}
                        >
                            {generating && currentFormat === "word" ? "Generating Word..." : "Generate Word"}
                        </button>
                    </div>
                </div>
            </form>
        </main>
    );
}

const styles = {
    label: {
        display: "block",
        fontSize: "1rem",
        marginBottom: "0.5rem",
        fontWeight: "bold",
        color: "#333",
    },
    input: {
        width: "100%",
        padding: "0.75rem",
        borderRadius: "4px",
        border: "1px solid #ddd",
        fontSize: "1rem",
    },
    select: {
        width: "100%",
        padding: "0.75rem",
        borderRadius: "4px",
        border: "1px solid #ddd",
        fontSize: "1rem",
    },
    button: {
        flex: 1,
        padding: "0.75rem",
        fontSize: "1rem",
        border: "none",
        borderRadius: "4px",
        backgroundColor: "#007bff",
        color: "#fff",
        cursor: "pointer",
        transition: "background-color 0.3s ease",
    },
    buttonHover: {
        backgroundColor: "#0056b3",
    },
};
