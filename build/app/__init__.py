from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import logging

db = SQLAlchemy()

def create_tables():
    from app import models
    db.create_all()  # Cria as tabelas no banco de dados
    logging.info("Tabelas criadas com sucesso")

def create_app():
    app = Flask(__name__)
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///notas.db'
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['SECRET_KEY'] = 'IFPB_SS_COPED'

    db.init_app(app)

    # Configurar logging
    logging.basicConfig(level=logging.INFO)

    with app.app_context():
        create_tables()

    # Importar as rotas
    from app.routes import main
    app.register_blueprint(main)

    return app
