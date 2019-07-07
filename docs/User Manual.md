## Installation
1. Download and Install [Anaconda](https://www.anaconda.com/distribution/#download-section)
	> NOTE: Don't forget to tick the command terminal option (which is NOT recommended)
2. Install packages
	```markdown
	python3
	pandas
	pywin32
	xlwings
	```
3. Enable macro settings in MS Excel.
4. Integrate `xlwings` with Excel. It would appear in separate tab inside Excel application.

## Cells
* #### Input:
	- `H5`: Company name
	- `I5`: Company location (if many)
	- `H20`: Contact name

* #### Output:
	- `B10`: Contact name
	- `B11`: Company name
	- `B12`: Company's location
	- `B13`: Company's address
	- `B14`: Company's phone no.
	- `D13`: Company's shipping address 

	> NOTE: The cell would be blank, if there is no corresponding data available in Excel sheet - `Customers`

## Testing
1. #### Search by Company
	- press <kbd>RESET</kbd> button to clear all existing values before making any search [RECOMMENDED]. 
	- type nothing and press <kbd>RUN</kbd> button. Then, see the message.
	- type "ALBERTA INFRASTRUCTURE" and press <kbd>RUN</kbd> button.
	- type "TECHMATION ELECTRIC & CONTROLS" and press <kbd>RUN</kbd> button. Then, choose any location out of the drop-down list.
2. #### Search by Contact
	- press <kbd>RESET</kbd> button to clear all existing values before making any search [RECOMMENDED]. 
	- type nothing and press <kbd>RUN</kbd> button. Then, see the message.
	- type "DAN WYATT" and press <kbd>RUN</kbd> button.

## Performance
* When any button is pressed for 1st time, it takes time to load the libraries (as opened for 1st time). So, it takes time of 5-10 seconds.
* But, from next run onwards it executes within 1 sec of time.
* Code is optimized currently.