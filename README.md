<h1>Scheduler Pro</h1>

<p>
This project is a WPF desktop application for managing and generating staff schedules. 
It allows users to combine manual availability entries with imported data from CSV files 
or LettuceMeet polls, then automatically generates optimized schedules based on configurable 
parameters. The application features an interactive grid interface for visual schedule 
editing, batch management, and PowerPoint export capabilities.
</p>

<p>
The application follows a multi-page workflow approach, guiding users through setup, 
data import, schedule generation, and editing phases with a modern, intuitive interface.
</p>

<hr>

<h2>Features</h2>
<ul>
    <li>Manual availability entry with interactive grid interface and drag-select functionality</li>
    <li>Multi-source data import from CSV files and LettuceMeet polls</li>
    <li>Automated schedule generation with configurable parameters</li>
    <li>Batch management for organizing groups of availability data</li>
    <li>Interactive schedule editing with add/remove row and column functionality</li>
    <li>PowerPoint export for presentation-ready schedule outputs</li>
    <li>Modern UI with card-based design and intuitive navigation</li>
    <li>Storage for saving and loading data between sessions</li>
</ul>

<hr>

<h2>Requirements</h2>
<ul>
    <li>.NET Framework 4.7.2 or later</li>
    <li>Windows 7 or later (WPF application)</li>
    <li>For LettuceMeet import: Internet connection and valid LettuceMeet event URL</li>
    <li>For PowerPoint export: Microsoft PowerPoint or compatible viewer</li>
</ul>

<hr>

<h2>Project Structure</h2>

<pre><code>
Views/                   # UI components and pages
Services/                # Core services: data import, export, scheduling algorithms
MainWindow.xaml          # Main application window with navigation and UI layout
MainWindow.xaml.cs       # Application logic and event handlers
SchedulerApp.csproj
App.xaml
App.xaml.cs
App.config
packages.config
readme.md
</code></pre>

<hr>

<h2>How It Works</h2>

<p>The application follows a sequential workflow:</p>

<h3>1. Setup Phase</h3>
<ul>
    <li>Configure basic scheduling parameters: time ranges, shift lengths, people per shift</li>
    <li>Select or create a batch of availability data</li>
    <li>Set weekend inclusion options</li>
</ul>

<h3>2. Data Collection Phase</h3>
<ul>
    <li><strong>Manual Entry</strong>: Use the interactive calendar and time grid to mark availability</li>
    <li><strong>CSV Import</strong>: Upload formatted CSV files with availability data</li>
    <li><strong>LettuceMeet Import</strong>: Direct integration with LettuceMeet polls to extract participant availability</li>
</ul>

<h3>3. Batch Management</h3>
<ul>
    <li>Combine availability data from multiple sources</li>
    <li>Save groups of availability as reusable batches</li>
    <li>Edit, rename, or delete batches as needed</li>
</ul>

<h3>4. Schedule Generation</h3>
<ul>
    <li>Select a batch and configure scheduling parameters</li>
    <li>Automated algorithm matches availability with shift requirements</li>
    <li>Generates optimized schedule with visual feedback on coverage</li>
</ul>

<h3>5. Schedule Editing</h3>
<ul>
    <li>Interactive grid interface for manual adjustments</li>
    <li>Add/remove rows (time intervals) and columns (days)</li>
    <li>Modify individual cell assignments</li>
    <li>Real-time statistics on coverage and assignments</li>
</ul>

<h3>6. Export</h3>
<ul>
    <li>Export schedules to PowerPoint presentations</li>
    <li>Structured output suitable for sharing and presentation</li>
</ul>

<hr>

<h2>Scheduling Algorithm</h2>

<p>The scheduling system uses a multi-phase optimization algorithm that balances staff availability with shift requirements while promoting fair workload distribution and shift continuity. The algorithm operates in several strategic phases:</p>

<h3>Phase 1: Critical Shift Assignment</h3>
<ul>
    <li>Identifies shifts with limited candidate availability (≤ required people)</li>
    <li>Prioritizes filling these shifts first to prevent coverage gaps</li>
    <li>Uses "hit map" analysis to identify shifts with 0-3 available candidates</li>
    <li>Assigns candidates based on their total available hours (prioritizing those with less availability)</li>
</ul>

<h3>Phase 2: Shift Continuity Enforcement</h3>
<ul>
    <li>Promotes back-to-back shift assignments within the same day</li>
    <li>Avoids creating split shifts (gaps between shifts for the same person)</li>
    <li>Prioritizes assigning people to consecutive shifts when possible</li>
    <li>Special attention to last shifts of the day to enable shift stacking</li>
</ul>

<h3>Phase 3: Remaining Shift Filling</h3>
<ul>
    <li>Processes remaining shifts ordered by candidate availability</li>
    <li>Uses a fairness scoring system that considers:
        <ul>
            <li>Assigned hours to date</li>
            <li>Availability utilization percentage</li>
            <li>Weekly hours distribution</li>
        </ul>
    </li>
    <li>Penalizes assignments that would create split shifts</li>
    <li>Gives priority to final shifts of the day to enable shift stacking from previous shifts</li>
</ul>

<h3>Phase 4: Dynamic Hour Balancing</h3>
<ul>
    <li>Monitors weekly hours across all staff members</li>
    <li>Identifies disparities in assigned hours</li>
    <li>Attempts to transfer shifts from over-assigned to under-assigned staff</li>
    <li>Only transfers when it won't create coverage issues or split shifts</li>
    <li>Iteratively balances workload towards the current maximum hours</li>
</ul>

<h3>Phase 5: Final Optimization</h3>
<ul>
    <li>Cleans up any remaining split shifts</li>
    <li>Removes earlier shifts from split-shift days when possible</li>
    <li>Ensures all assignments respect availability constraints</li>
</ul>

<h3>Algorithm Features</h3>
<ul>
    <li><strong>Constraint-based scheduling</strong>: Respects all configured parameters (shift length, people per shift, time ranges)</li>
    <li><strong>Availability matching</strong>: Only assigns staff to shifts they marked as available</li>
    <li><strong>Shift stacking preference</strong>: Encourages consecutive shifts over split shifts</li>
    <li><strong>Workload fairness</strong>: Distributes hours evenly across available staff</li>
    <li><strong>Real-time optimization</strong>: Multiple phases refine the schedule progressively</li>
    <li><strong>Coverage maximization</strong>: Prioritizes filling understaffed shifts</li>
</ul>

<h3>Data Structures Used</h3>
<ul>
    <li><strong>PersonAvailability</strong>: Individual availability slots with date and time ranges</li>
    <li><strong>ScheduleConfig</strong>: Scheduling parameters (opening/closing times, shift length, people per shift)</li>
    <li><strong>Shift</strong>: Individual shift instances with assigned people and requirements</li>
    <li><strong>HitMap</strong>: Candidate analysis for each shift to identify critical coverage needs</li>
    <li><strong>Tracking dictionaries</strong>: Monitor assigned hours, weekly hours, and shift continuity per person</li>
</ul>

<hr>

<h2>Notes</h2>

<p>
This application was developed as a personal project for schedule management needs. 
It is provided free of charge for non-commercial and non-profit use. While the core 
functionality is complete and operational, the project should be considered feature-complete 
rather than actively maintained.
</p>

<p>
Users should ensure compliance with any terms of service when importing data from 
external services like LettuceMeet. The application is designed for Windows desktop 
environments and requires .NET Framework.
</p>
