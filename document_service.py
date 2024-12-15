import os
import uno
from flask import Flask, request, jsonify

app = Flask(__name__)

def connect_to_libreoffice():
    """Establish connection to LibreOffice in headless mode."""
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    context = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
    return context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

@app.route("/create_document", methods=["POST"])
def create_document():
    """Create a document in LibreOffice."""
    try:
        desktop = connect_to_libreoffice()
        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        text = doc.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, request.json.get("content", "empty"), 0)
    
        # Save the document
        output_path = request.json.get("output_path", "output.odt")

        output_path = output_path.replace('\\', '/')

        output_path = "file:///{file_path}".format(file_path=output_path)
        doc.storeToURL(output_path, ())
        doc.close(True)

        return jsonify({"message": "Document created", "path": output_path})
    except Exception as e:
        return jsonify({"erroror": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
