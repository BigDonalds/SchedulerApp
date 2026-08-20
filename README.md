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
    <li>PowerPoint export with multiple design templates for presentation-ready outputs</li>
    <li>Animated background themes that change based on time of day</li>
    <li>Modern UI with card-based design and intuitive navigation</li>
    <li>Persistent storage for saving and loading data between sessions</li>
</ul>

<hr>

<h2>Requirements</h2>
<ul>
    <li>.NET Framework 4.8.1</li>
    <li>Windows 7 or later (WPF application)</li>
    <li>For LettuceMeet import: Internet connection and valid LettuceMeet event URL</li>
    <li>For PowerPoint export: Microsoft PowerPoint or compatible viewer</li>
</ul>

<hr>

<h2>Project Structure</h2>

<pre><code>
Views/                   # UI components and pages
Services/                # Core services: scheduling, export, import, storage, animations
Styles/                  # XAML resource dictionaries for UI theming
Templates/               # PowerPoint export templates
Icons/                   # Application icons and template preview images
Properties/              # Resources and settings
MainWindow.xaml          # Main application window with navigation and UI layout
MainWindow.xaml.cs       # Application logic and event handlers
SchedulerApp.csproj      # Project configuration
App.xaml
App.xaml.cs
App.config
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
    <li>Multiple design templates with structured output suitable for sharing and presentation</li>
</ul>

<hr>

<h2>Scheduling Algorithm</h2>

<p>The scheduling system uses a 13-phase optimization algorithm that balances staff availability with shift requirements while promoting fair workload distribution, shift continuity, and full coverage. The algorithm operates in the following phases:</p>

<h3>Phase 0: Pre-Identify Last-Shift Stacking Candidates</h3>
<ul>
    <li>Before any scheduling, examines every day to find days where the last shift is shorter than the configured shift length</li>
    <li>Identifies people available for BOTH of the last two shifts of the day</li>
    <li>Locks those two shifts so other phases cannot modify them</li>
    <li>Guarantees continuity for the final short shift of each day</li>
</ul>

<h3>Phase 1: Tiered Heatmap Assignment</h3>
<ul>
    <li>Builds a "hit map" for every shift showing how many candidates are available</li>
    <li>Assigns shifts progressively in tiers: 1x shift length, 2x, 3x, and so on up to 8 tiers</li>
    <li>Within each tier, processes shifts by heatmap priority (fewest candidates first)</li>
    <li>Caps each person's assigned hours at the current tier target to ensure fair distribution early on</li>
    <li>Also handles partial availability assignments where candidates cover only part of a shift</li>
</ul>

<h3>Phase 2: Last Short Shift Stacking</h3>
<ul>
    <li>For days where the final shift is shorter than normal, prioritizes people from the previous shift to fill it</li>
    <li>Falls back to any available candidate if the previous shift's staff cannot cover it</li>
</ul>

<h3>Phase 3: Partial Availability Overlaps</h3>
<ul>
    <li>For remaining unfilled shifts, finds candidates with partial time overlap (at least 30 minutes)</li>
    <li>Assigns them with start/end time annotations so coverage is tracked accurately</li>
    <li>Skips assignments that would create same-day gaps</li>
</ul>

<h3>Phase 4: Remaining Shift Filling with Daily Cap Penalty</h3>
<ul>
    <li>Fills remaining shifts ordered by candidate scarcity</li>
    <li>Uses a fairness scoring system that considers:
        <ul>
            <li>Assigned hours to date</li>
            <li>Availability utilization percentage</li>
            <li>Weekly hours distribution</li>
            <li>Progressive daily cap penalty (quadratic scaling as daily hours approach the cap)</li>
            <li>Consecutive shift accumulation penalty</li>
        </ul>
    </li>
    <li>Relaxes split-shift restrictions when not enough candidates are available</li>
</ul>

<h3>Phase 5: Low-Hour Extra Assignments</h3>
<ul>
    <li>Identifies people with less than 4 hours of total weekly availability</li>
    <li>Adds them as one extra person on already-staffed shifts they partially overlap</li>
    <li>Respects daily hour caps and same-day gap prevention</li>
</ul>

<h3>Phase 6: Same-Day Gap Prevention Sweep</h3>
<ul>
    <li>Scans every person's shift assignments per day for gaps between shifts</li>
    <li>Removes people from shifts that create non-adjacent gaps (when staffing allows)</li>
    <li>Preferentially removes from the earlier shift when possible</li>
</ul>

<h3>Phase 7: Smart Last-Two-Shift Selection</h3>
<ul>
    <li>After all regular scheduling, selects the best candidates for the locked last-two-shift blocks</li>
    <li>Filters out candidates who would create same-day gaps</li>
    <li>Picks the candidates with the lowest weekly hours</li>
</ul>

<h3>Phase 8: Zero-Hour Pity Assignments</h3>
<ul>
    <li>Finds people with 0 assigned weekly hours</li>
    <li>Gives them up to 3 hours of work via three strategies:
        <ul>
            <li>One large chunk overlapping at least 80% of a full shift</li>
            <li>Consecutive shifts on the same day (e.g. 13:00-14:30 + 14:30-16:00)</li>
            <li>Smaller single partial shifts (minimum 30 minutes)</li>
        </ul>
    </li>
</ul>

<h3>Phase 9: Partial Coverage Gap Filler</h3>
<ul>
    <li>Detects gaps where partial-assignment people leave early and no one covers the remaining time</li>
    <li>First tries to extend people from the next shift into the gap</li>
    <li>Falls back to any available candidate for the gap window</li>
</ul>

<h3>Phase 10: Half-Hour Understaffing Scan & Fix</h3>
<ul>
    <li>Walks every half-hour slot across the full schedule</li>
    <li>Counts ACTUAL presence (partial people only count for the time they are really working)</li>
    <li>Finds contiguous understaffed windows and attempts to fix them with available candidates</li>
    <li>Generates detailed understaffing alerts with candidate suggestions for any unresolved windows</li>
</ul>

<h3>Phase 11: Overwork/Underwork Balance Exchange</h3>
<ul>
    <li>Calculates the weekly average assigned hours across all staff</li>
    <li>Transfers entire stacked blocks of shifts from overworked to underworked people</li>
    <li>Handles full+partial, partial+full, and partial+partial block combinations together</li>
    <li>Uses a gradual balance guard so no recipient exceeds the weekly average</li>
    <li>Second pass allows transfers with backfill candidates when removal would understaff a shift</li>
</ul>

<h3>Phase 12: Final Understaffing Re-Scan</h3>
<ul>
    <li>Re-runs the half-hour understaffing scan after workload balancing</li>
    <li>Catches any newly-created coverage gaps from transfers</li>
    <li>Ensures the final alerts reflect the true schedule state</li>
</ul>

<h3>Algorithm Rules</h3>
<ul>
    <li><strong>Rule 1 - Heatmap priority</strong>: Shifts with fewer candidates are filled first</li>
    <li><strong>Rule 2 - Tiered fair distribution</strong>: Hours are balanced progressively across tiers</li>
    <li><strong>Rule 3 - Last-two-shift stacking</strong>: Locked blocks ensure the final short shift is covered by continuity</li>
    <li><strong>Rule 4 - Daily cap penalty</strong>: Quadratic penalty prevents daily overload (45% max ratio, 30% penalty threshold)</li>
    <li><strong>Rule 5 - Same-day gap prevention</strong>: No gaps between a person's shifts on the same day</li>
    <li><strong>Rule 6 - Partial availability</strong>: People with partial overlap can be assigned with precise start/end times</li>
    <li><strong>Rule 7 - Low-hour extras</strong>: People with minimal availability still get included</li>
</ul>

<h3>Data Structures Used</h3>
<ul>
    <li><strong>PersonAvailability</strong>: Individual availability slots with date and time ranges</li>
    <li><strong>ScheduleConfig</strong>: Scheduling parameters (opening/closing times, shift length, people per shift, closed days, stacking toggle)</li>
    <li><strong>Shift</strong>: Individual shift instances with assigned people, position in day, first/last flags, and lock state</li>
    <li><strong>PersonAssignment</strong>: Detailed assignment with optional start/end overrides for partial coverage</li>
    <li><strong>HitMap</strong>: Candidate analysis for each shift to identify critical coverage needs</li>
    <li><strong>UnderstaffingAlert</strong>: Reports unresolved coverage windows with suggested candidates</li>
    <li><strong>Tracking dictionaries</strong>: Monitor assigned hours, weekly hours, daily caps, consecutive shifts, and shift continuity per person</li>
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
