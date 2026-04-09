Problem: Manual collection of match data (CDN IDs, MGG IDs, start times) from multiple source sheets was consuming ~40 man-hours per week and was prone to human error.

Solution: A custom Google Apps Script that scans multiple data sources, performs flexible header detection (fuzzy matching for naming conventions), filters by date range, and aggregates everything into a standardized, "ready-to-use" format for the operations team.

Key Features:
* Smart Header Mapping: Automatically identifies columns regardless of language (UA/RU/EN) or minor naming variations.

Automated Filtering: Dynamic date-windowing (current day + 7-day outlook).

Data Validation: Built-in parsing for complex time/date strings to ensure aggregate integrity.
* Efficiency: Reduced the weekly manual workload by 40 hours.
