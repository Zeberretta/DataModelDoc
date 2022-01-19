venv/bin/activate: requirements.txt
	python3 -m venv venv
	./venv/bin/python3 -m pip install --upgrade pip
	./venv/bin/pip3 install -r requirements.txt


run: venv/bin/activate
	./venv/bin/python3 GUI.py	

clean:
	rm -rf __pycache__
	rm -rf venv
