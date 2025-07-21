from app import db


    
class Aluno(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    matricula = db.Column(db.String(10), nullable=False)
    nome = db.Column(db.String(255), nullable=False)
    curso = db.Column(db.String(255), nullable=False)
    ano = db.Column(db.Integer, nullable=False)
    disciplina = db.Column(db.String(255), nullable=False)
    media1 = db.Column(db.Integer, nullable=False, default=0)
    frequencia1 = db.Column(db.Integer, nullable=False, default=0)
    media2 = db.Column(db.Integer, nullable=False, default=0)
    frequencia2 = db.Column(db.Integer, nullable=False, default=0)
    media3 = db.Column(db.Integer, nullable=False, default=0)
    frequencia3 = db.Column(db.Integer, nullable=False, default=0)
    media4 = db.Column(db.Integer, nullable=False, default=0)
    frequencia4 = db.Column(db.Integer, nullable=False, default=0)
    final = db.Column(db.Integer, nullable=False, default=0)

    def __repr__(self):
        return f'<Aluno {self.nome}>'