# Excel 2 Graph

**Excel 2 Graph** is a user-friendly desktop application designed to convert your Excel data into visually appealing graphs. Whether you're a data analyst, researcher, or business professional, this tool simplifies the process of generating customizable graphs from your Excel spreadsheets.

## Features

- **Easy Folder Selection:** Choose input folders containing `.xlsx` files and designate output folders for saving generated graphs.
- **Customizable Data Series:** Configure the number of data series, assign colors and markers, and rename series for better clarity.
- **Real-Time Preview:** View a live preview of your graph as you adjust settings, ensuring your final output meets your expectations.
- **Flexible Graph Settings:** Customize x and y-axis labels, legend positions, and graph fonts to match your presentation needs.
- **Interactive UI:** Intuitive interface built with `customtkinter` for a seamless user experience.
- **Automated Graph Generation:** Convert multiple Excel files to graphs with a single click, saving you time and effort.

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/berkay-digital/excel2graph.git
   cd excel2graph
   ```

2. **Create a Virtual Environment (Optional but Recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare Your Excel Files:**
   - Place your `.xlsx` files in the `./excel` directory. Ensure that you use the correct sheet names for each file.

2. **Run the Application:**
   ```bash
   python main.py
   ```

3. **Configure Settings:**
   - **Input Folder:** Select the folder containing your Excel files.
   - **Output Folder:** Choose where the generated graphs will be saved.
   - **Number of Data Series:** Adjust the number of data series to display.
   - **Series Configuration:** Assign colors and markers to each data series and rename them as needed.
   - **Axis Labels:** Set custom labels for the x and y axes.
   - **Legend Position:** Choose where the legend appears on the graph.
   - **Visibility Options:** Toggle the visibility of the legend and data symbols.
   - **Font Selection:** Select your preferred font for graph labels and titles.

4. **Generate Graphs:**
   - Click the **"Create Graphs"** button to convert your Excel data into graphs. The application will process each Excel file and save the corresponding graph as a `.png` image in the output folder.

## Dependencies

- [customtkinter](https://github.com/TomSchimansky/CustomTkinter) == 5.2.0
- [matplotlib](https://matplotlib.org/) == 3.8.2
- [numpy](https://numpy.org/) == 1.25.2
- [pandas](https://pandas.pydata.org/) == 2.2.3
- [scipy](https://www.scipy.org/) == 1.14.1

All dependencies are listed in the `requirements.txt` file and can be installed using `pip`.

## Contributing

Contributions are welcome! If you'd like to enhance the application or fix any issues, please follow these steps:

## License

This project is licensed under the [MIT License](LICENSE).

## Contact

Developed by [Berkay](https://berkay.digital). For any inquiries or feedback, please reach out through the [GitHub Issues](https://github.com/berkay-digital/excel2graph/issues) page.
