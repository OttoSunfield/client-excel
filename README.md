# Realtime quotes in Excel

This document describes how to get streaming quotes in Microsoft Excel.

## Installation

1. Install Script Lab in Excel, as described [here](https://github.com/OfficeDev/script-lab/blob/master/README.md#top).

2. Open Script Lab in Excel and in the code tab select "Import".

3. Paste the YAML snippet in this repository. This will create the sample of using the Binck API for realtime quotes.

## Usage

The Script Lab snippet requires an access token and an account number. The access token is used for authorization, and gives access to the instruments available for your account.

1. Create an access token using [this website](http://www.basement.nl/binck/demo.html).

2. Open Excel, select the Add-in Script Lab.

3. Add the token to the HTML of the snippet (or insert it later, after running the snippet).

4. Do the same for the account number.

5. Run the snippet from the ribbon menu. Script Lab will compile the code and show token and account number edits, with three buttons, "Get indices", "Get streaming quote updates" and "Stop streaming updates".

6. Click the button "Get indices". Excel will use the Bearer token to connect with the Binck API and if connection can be established, it will request the instruments list with the indices.

7. Click "Get streaming quote updates". This will connect to the Binck streamer to setup a websocket connection with realtime updates.

8. When done, click "Stop streaming quotes".

## Idea's for improvement

The instruments in this example are lists, but it can be the positions as well, or any other instrument.
The quotes shown in Excel are not formatted. Color is not there as well. This is just a quick start.

Tokens are valid for one hour. You can extend this, but this requires a backend. See instructions in the [Javascript showcase](https://github.com/binckbank-api/client-js).
