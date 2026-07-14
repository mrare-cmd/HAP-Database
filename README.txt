HAP DATABASE - README
=====================

OPEN IT
-------
   Double-click:  Open HAP Database.bat
   Your browser opens with the whole database - sorted, filterable, like Excel.
   (A small black window stays open behind it. That's the engine - just leave
    it be, and close it when you're finished.)

USE IT
------
   - Type under any column header to filter (e.g. County = Los Angeles).
   - Click a header to sort (e.g. Rent-to-FMR Ratio).
   - "Export to Excel" saves what you're looking at, formatted like the old
     January master table.

GET THE CURRENT MONTH
---------------------
   Click the green  "Update to current month"  button, top-left.
   It pulls this month's HUD files, rebuilds everything, and the "Data as of"
   date changes to today. Takes about 5 minutes. That's the only button you
   ever need for updates.

   (First time only: it quietly installs a few components and, if Python isn't
    on the PC, opens the download page - install it, tick "Add to PATH", done.)

WHY A LITTLE ENGINE INSTEAD OF JUST A WEBPAGE
---------------------------------------------
   Downloading from HUD and saving files are things a plain webpage isn't
   allowed to do, for security. The little engine does that part so the button
   in the page can work.

THE FIX
-------
   The old automation summed duplicate rent rows (740 S. Olive showed $43,144
   instead of $2,911). Fixed and verified against your January database:
   99.6% exact match.
