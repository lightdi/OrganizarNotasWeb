import os
from flask import Blueprint, render_template, request, redirect, url_for, flash
from app import db
import re
import pandas as pd
from app.models import  Aluno
from werkzeug.utils import secure_filename

main = Blueprint('main', __name__)

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Garantir que a pasta de uploads exista
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    """Verifica se o arquivo tem extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@main.route('/')
def index():
    # Consulta todos os cursos distintos dos alunos
    cursos = db.session.query(Aluno.curso,Aluno.ano).distinct().all()

    return render_template('index.html', cursos=cursos)

@main.route('/curso/<curso_nome>/<ano>')
def curso(curso_nome, ano):
    # Consulta os alunos do curso e ano selecionados
    alunos = Aluno.query.filter_by(curso=curso_nome, ano=ano).order_by( Aluno.nome, Aluno.disciplina).all()

    return render_template('curso.html', curso_nome=curso_nome, ano=ano, alunos=alunos)

@main.route('/upload', methods=['GET', 'POST'])
def upload_file():
    """Rota para upload e inserção/atualização de dados"""
    print("Entrou")
    if request.method == 'POST':
        if 'file' not in request.files:
            flash("Nenhum arquivo enviado.", "danger")
            print("Nenhum arquivo enviado.")
            return redirect('/upload') 

        file = request.files['file']
        bimestre = request.form.get('bimestre')
        ano = request.form.get('ano')
       
        if file.filename == '':
            flash("Nenhum arquivo selecionado.", "danger")
            print("Nenhum arquivo selecionado.")
            return redirect('/upload') 

        if not bimestre or bimestre not in ["1", "2", "3", "4"]:
            flash("Selecione um bimestre válido.", "danger")
            print("Selecione um bimestre válido.")
            return redirect('/upload') 
        
        if not ano or ano not in ["1", "2", "3"]:
            flash("Selecione um ano válido.", "danger")
            return redirect('/upload')

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
           
            try:
                dfs = pd.read_excel(filepath, sheet_name=None)
                df = dfs.get('Planilha1')
                print("Carregou")
                if df is None:
                    flash("Erro: Planilha 'Planilha1' não encontrada no arquivo.", "danger")
                    print("Erro: Planilha 'Planilha1' não encontrada no arquivo.")
                    return redirect('/upload')

                novos = 0
                atualizados = 0
                for i, linha in df.iterrows():
                    if i == 191:
                        print(linha)
                    matricula = str(linha["Matricula"])
                    nome = linha["Nome"]
                    disciplina = linha['Disciplina']
                    curso = linha['Curso']
                    nota = linha['Nota']
                    frequencia = linha['Frequencia']
                    media = extrair_media(nota)

                    aluno = Aluno.query.filter_by(matricula=matricula, disciplina=disciplina).first()

                    if aluno:
                        # Atualiza os dados do bimestre correto
                        if bimestre == "1":
                            aluno.media1 = int(media)
                            aluno.frequencia1 = int(frequencia)
                        elif bimestre == "2":
                            aluno.media2 = int(media)
                            aluno.frequencia2 = int(frequencia)
                        elif bimestre == "3":
                            aluno.media3 = int(media)
                            aluno.frequencia3 = int(frequencia)
                        elif bimestre == "4":
                            aluno.media4 = int(media)
                            aluno.frequencia4 = int(frequencia)
                        atualizados += 1
                    else:
                        # Insere um novo aluno
                        aluno = Aluno(
                            matricula=matricula,
                            nome=nome,
                            curso=curso,
                            ano = ano,
                            serie = int(ano),
                            disciplina=disciplina,
                        )
                        if bimestre == "1":
                            aluno.media1 = int(media)
                            aluno.frequencia1 = int(frequencia)
                        elif bimestre == "2":
                            aluno.media2 = int(media)
                            aluno.frequencia2 = int(frequencia)
                        elif bimestre == "3":
                            aluno.media3 = int(media)
                            aluno.frequencia3 = int(frequencia)
                        elif bimestre == "4":
                            aluno.media4 = int(media)
                            aluno.frequencia4 = int(frequencia)

                        db.session.add(aluno)
                        novos += 1
                        aluno = None

                db.session.commit()
                flash(f"Importação concluída! {novos} novos alunos inseridos, {atualizados} atualizados.", "success")

            except Exception as e:
                db.session.rollback()
                print(e)
                flash(f"Erro ao processar o arquivo: {e}", "danger")

            return redirect('/')

    return render_template('upload.html')

def extrair_media(nota):
    media = None
    if pd.notna(nota):
        # Atualiza o padrão para capturar apenas a média
        pattern = r'Media:(\d+|[-])'
        matches = re.findall(pattern, nota)

        if matches:
            media = matches[0]

            # Converte '-' para 0, caso a média seja '-'
            media = 0 if media == '-' else int(media)

    return media