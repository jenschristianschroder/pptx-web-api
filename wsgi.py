from flask import Flask

def create_app():
    app = Flask(__name__)

    # Load configuration from environment variables or a config file
    app.config.from_mapping(
        # Add your configuration settings here
    )

    # Register blueprints or routes
    from app.routes import main as main_blueprint
    app.register_blueprint(main_blueprint)

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=8000, debug=True)