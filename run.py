from flask import Flask, render_template
app = Flask(__name__)

#from automate_test import open_and_ID
#
@app.route("/")
def template_test():
    return render_template('gui_home.html')


if __name__ == '__main__':
    app.run(debug=True)
#
# print(Flask(__name__))