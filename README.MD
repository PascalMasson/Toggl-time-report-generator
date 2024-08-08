Got it! Here’s the updated README reflecting that the script is in `index.js`:

---

# Toggl Time Report Generator

This script generates a time report from Toggl time entries and optionally sends it via email. It uses several packages to fetch time entries, process the data, create an Excel report, and handle email sending.

## Features

- **Generate reports for the current week, last week, or a specified week.**
- **Format and summarize Toggl time entries.**
- **Create and save a report in Excel format.**
- **Send the generated report via email.**

## Requirements

- Node.js (v14 or later)
- npm (Node Package Manager)

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/your-repo/toggl-time-report-generator.git
    ```

2. Navigate to the project directory:

    ```bash
    cd toggl-time-report-generator
    ```

3. Install dependencies:

    ```bash
    npm install
    ```

## Configuration

Create a `.env` file in the root directory of the project with the following environment variables:

```env
TOGGL_TRACK_API_TOKEN=your_toggl_api_token
EMAIL_HOST=your_email_host
EMAIL_PORT=your_email_port
EMAIL_USER=your_email_address
EMAIL_PASS=your_email_password
EMAIL_TO=recipient_email_address
EMAIL_CC_SELF=optional_cc_email_address
```

## Usage

You can run the script in two ways:

- **Using `node`:**

    Run the script with one of the following options:

    - **Generate a report for the current week:**

        ```bash
        node index.js --current
        ```

    - **Generate a report for the last week:**

        ```bash
        node index.js --last
        ```

    - **Generate a report for a specific week (format: YYYY-MM-DD):**

        ```bash
        node index.js --week YYYY-MM-DD
        ```

    For example, to generate a report for the week starting on July 1, 2024:

    ```bash
    node index.js --week 2024-07-01
    ```

- **Using `npm run start`:**

    You can also run the script using npm:

    ```bash
    npm run start -- --current
    ```

    ```bash
    npm run start -- --last
    ```

    ```bash
    npm run start -- --week YYYY-MM-DD
    ```

    **Note:** The double dash `--` separates the npm script arguments from the arguments passed to the script itself.

## Emailing the Report

After generating the report, the script will prompt whether you want to send the file via email. Type `yes` to send the report, or `no` to skip this step.

## Notes

- Make sure you have an appropriate API token from Toggl.
- Ensure the email configuration in the `.env` file is correct and that the SMTP server details are accurate.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

Feel free to update any other details or sections as needed!