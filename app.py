from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta, date
import pandas as pd
import os
import io

app = Flask(__name__)
app.config['SECRET_KEY'] = 'sistema-frota-2025-secret-key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///sistema_frota.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Por favor, faça login para acessar esta página.'

# Modelos do Banco de Dados
class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    nome = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    ativo = db.Column(db.Boolean, default=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class Colaborador(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    matricula = db.Column(db.String(50), unique=True, nullable=False)
    cpf = db.Column(db.String(14), unique=True)
    telefone = db.Column(db.String(20))
    email = db.Column(db.String(120))
    veiculo_vinculado = db.Column(db.String(20))
    ativo = db.Column(db.Boolean, default=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relacionamentos
    pontos = db.relationship('Ponto', backref='colaborador', lazy=True, cascade='all, delete-orphan')
    frotas = db.relationship('Frota', backref='motorista_obj', lazy=True)
    descontos = db.relationship('Desconto', backref='colaborador', lazy=True)

class Ponto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    colaborador_id = db.Column(db.Integer, db.ForeignKey('colaborador.id'), nullable=False)
    data_hora = db.Column(db.DateTime, nullable=False)
    tipo = db.Column(db.String(10), nullable=False)  # entrada ou saida
    observacao = db.Column(db.Text)
    extraordinario = db.Column(db.Boolean, default=False)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class Frota(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.Date, nullable=False)
    veiculo = db.Column(db.String(20), nullable=False)
    motorista_id = db.Column(db.Integer, db.ForeignKey('colaborador.id'), nullable=False)
    hora_saida = db.Column(db.Time)
    hora_retorno = db.Column(db.Time)
    km_inicial = db.Column(db.Float)
    km_final = db.Column(db.Float)
    observacao = db.Column(db.Text)
    status = db.Column(db.String(20), default='conforme')  # conforme ou extraordinaria
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class Desconto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    colaborador_id = db.Column(db.Integer, db.ForeignKey('colaborador.id'), nullable=False)
    data = db.Column(db.Date, nullable=False)
    motivo = db.Column(db.Text, nullable=False)
    valor = db.Column(db.Float, nullable=False)
    status = db.Column(db.String(20), default='pendente')  # pendente, aprovado, cancelado
    frota_id = db.Column(db.Integer, db.ForeignKey('frota.id'))
    automatico = db.Column(db.Boolean, default=False)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    
    frota = db.relationship('Frota', backref='descontos')

@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))

def is_admin():
    """Verifica se o usuário logado é o administrador."""
    return current_user.is_authenticated and current_user.username == 'admin'

# Funções auxiliares
def verificar_ponto_motorista(motorista_id, data):
    """Verifica se o motorista tem ponto registrado na data"""
    inicio = datetime.combine(data, datetime.min.time())
    fim = datetime.combine(data, datetime.max.time())
    
    ponto = Ponto.query.filter(
        Ponto.colaborador_id == motorista_id,
        Ponto.data_hora >= inicio,
        Ponto.data_hora <= fim
    ).first()
    
    return ponto is not None

def gerar_desconto_automatico(frota_registro):
    """Gera desconto automático para motorista sem ponto"""
    # Verifica se já existe desconto para este registro
    desconto_existente = Desconto.query.filter_by(
        frota_id=frota_registro.id,
        automatico=True
    ).first()
    
    if not desconto_existente and not verificar_ponto_motorista(frota_registro.motorista_id, frota_registro.data):
        # Calcula o valor do desconto (exemplo: R$ 50,00 por falta de registro)
        valor_desconto = 50.00
        
        # Se há quilometragem registrada, adiciona valor por km
        if frota_registro.km_inicial and frota_registro.km_final:
            km_rodado = frota_registro.km_final - frota_registro.km_inicial
            if km_rodado > 0:
                valor_desconto += km_rodado * 0.50  # R$ 0,50 por km
        
        desconto = Desconto(
            colaborador_id=frota_registro.motorista_id,
            data=frota_registro.data,
            motivo=f'Ausência de registro de ponto - Veículo {frota_registro.veiculo}',
            valor=valor_desconto,
            status='pendente',
            frota_id=frota_registro.id,
            automatico=True
        )
        
        db.session.add(desconto)
        db.session.commit()
        
        return desconto
    
    return None

# Rotas de Autenticação
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        user = Usuario.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password_hash, password):
            if user.ativo:
                login_user(user)
                flash('Login realizado com sucesso!', 'success')
                return redirect(url_for('index'))
            else:
                flash('Usuário inativo. Contate o administrador.', 'error')
        else:
            flash('Usuário ou senha incorretos.', 'error')
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Logout realizado com sucesso.', 'success')
    return redirect(url_for('login'))

# Rotas Principais
@app.route('/')
@login_required
def index():
    # Estatísticas do dashboard
    total_colaboradores = Colaborador.query.filter_by(ativo=True).count()
    total_pontos_hoje = Ponto.query.filter(
        db.func.date(Ponto.data_hora) == date.today()
    ).count()
    total_descontos_pendentes = Desconto.query.filter_by(status='pendente').count()
    total_frota_hoje = Frota.query.filter_by(data=date.today()).count()
    
    # Últimos registros
    ultimos_pontos = Ponto.query.order_by(Ponto.data_hora.desc()).limit(5).all()
    ultimos_descontos = Desconto.query.filter_by(status='pendente').order_by(Desconto.criado_em.desc()).limit(5).all()
    
    return render_template('index.html',
                         total_colaboradores=total_colaboradores,
                         total_pontos_hoje=total_pontos_hoje,
                         total_descontos_pendentes=total_descontos_pendentes,
                         total_frota_hoje=total_frota_hoje,
                         ultimos_pontos=ultimos_pontos,
                         ultimos_descontos=ultimos_descontos)

# Rotas de Colaboradores
@app.route('/colaboradores')
@login_required
def colaboradores():
    colaboradores = Colaborador.query.all()
    return render_template('colaboradores.html', colaboradores=colaboradores)

@app.route('/colaborador/novo', methods=['GET', 'POST'])
@login_required
def novo_colaborador():
    if request.method == 'POST':
        colaborador = Colaborador(
            nome=request.form['nome'],
            matricula=request.form['matricula'],
            cpf=request.form.get('cpf'),
            telefone=request.form.get('telefone'),
            email=request.form.get('email'),
            veiculo_vinculado=request.form.get('veiculo_vinculado'),
            ativo=True
        )
        
        try:
            db.session.add(colaborador)
            db.session.commit()
            flash('Colaborador cadastrado com sucesso!', 'success')
            return redirect(url_for('colaboradores'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao cadastrar colaborador: {str(e)}', 'error')
    
    return render_template('colaborador_form.html', colaborador=None)

@app.route('/colaborador/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_colaborador(id):
    colaborador = Colaborador.query.get_or_404(id)
    
    if request.method == 'POST':
        colaborador.nome = request.form['nome']
        colaborador.matricula = request.form['matricula']
        colaborador.cpf = request.form.get('cpf')
        colaborador.telefone = request.form.get('telefone')
        colaborador.email = request.form.get('email')
        colaborador.veiculo_vinculado = request.form.get('veiculo_vinculado')
        colaborador.ativo = 'ativo' in request.form
        
        try:
            db.session.commit()
            flash('Colaborador atualizado com sucesso!', 'success')
            return redirect(url_for('colaboradores'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar colaborador: {str(e)}', 'error')
    
    return render_template('colaborador_form.html', colaborador=colaborador)

@app.route('/colaborador/excluir/<int:id>')
@login_required
def excluir_colaborador(id):
    colaborador = Colaborador.query.get_or_404(id)
    
    try:
        db.session.delete(colaborador)
        db.session.commit()
        flash('Colaborador excluído com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir colaborador: {str(e)}', 'error')
    
    return redirect(url_for('colaboradores'))

# Rotas de Ponto
@app.route('/pontos')
@login_required
def pontos():
    pontos = Ponto.query.order_by(Ponto.data_hora.desc()).all()
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    return render_template('pontos.html', pontos=pontos, colaboradores=colaboradores)

@app.route('/ponto/novo', methods=['POST'])
@login_required
def novo_ponto():
    try:
        data = request.form['data']
        hora = request.form['hora']
        data_hora = datetime.strptime(f"{data} {hora}", "%Y-%m-%d %H:%M")
        
        ponto = Ponto(
            colaborador_id=request.form['colaborador_id'],
            data_hora=data_hora,
            tipo=request.form['tipo'],
            observacao=request.form.get('observacao', ''),
            extraordinario='extraordinario' in request.form
        )
        
        db.session.add(ponto)
        db.session.commit()
        
        flash('Ponto registrado com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao registrar ponto: {str(e)}', 'error')
    
    return redirect(url_for('pontos'))

@app.route('/ponto/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_ponto(id):
    ponto = Ponto.query.get_or_404(id)
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    
    if request.method == 'POST':
        try:
            data = request.form['data']
            hora = request.form['hora']
            ponto.colaborador_id = request.form['colaborador_id']
            ponto.data_hora = datetime.strptime(f"{data} {hora}", "%Y-%m-%d %H:%M")
            ponto.tipo = request.form['tipo']
            ponto.observacao = request.form.get('observacao', '')
            ponto.extraordinario = 'extraordinario' in request.form
            
            db.session.commit()
            flash('Ponto atualizado com sucesso!', 'success')
            return redirect(url_for('pontos'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar ponto: {str(e)}', 'error')
            
    return render_template('ponto_form.html', ponto=ponto, colaboradores=colaboradores)


@app.route('/ponto/excluir/<int:id>')
@login_required
def excluir_ponto(id):
    ponto = Ponto.query.get_or_404(id)
    
    try:
        db.session.delete(ponto)
        db.session.commit()
        flash('Ponto excluído com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir ponto: {str(e)}', 'error')
    
    return redirect(url_for('pontos'))

# Rotas de Frota
@app.route('/frota')
@login_required
def frota():
    registros = Frota.query.order_by(Frota.data.desc()).all()
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    return render_template('frota.html', registros=registros, colaboradores=colaboradores)

@app.route('/frota/novo', methods=['GET', 'POST'])
@login_required
def novo_frota():
    if request.method == 'POST':
        try:
            frota_registro = Frota(
                data=datetime.strptime(request.form['data'], '%Y-%m-%d').date(),
                veiculo=request.form['veiculo'],
                motorista_id=request.form['motorista_id'],
                hora_saida=datetime.strptime(request.form['hora_saida'], '%H:%M').time() if request.form.get('hora_saida') else None,
                hora_retorno=datetime.strptime(request.form['hora_retorno'], '%H:%M').time() if request.form.get('hora_retorno') else None,
                km_inicial=float(request.form['km_inicial']) if request.form.get('km_inicial') else None,
                km_final=float(request.form['km_final']) if request.form.get('km_final') else None,
                observacao=request.form.get('observacao', '')
            )
            
            # Verifica ponto e define status
            if not verificar_ponto_motorista(frota_registro.motorista_id, frota_registro.data):
                frota_registro.status = 'extraordinaria'
            
            db.session.add(frota_registro)
            db.session.commit()
            
            # Gera desconto automático se necessário
            if frota_registro.status == 'extraordinaria':
                desconto = gerar_desconto_automatico(frota_registro)
                if desconto:
                    flash('Desconto automático gerado por ausência de ponto!', 'warning')
            
            flash('Registro de frota criado com sucesso!', 'success')
            return redirect(url_for('frota'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao criar registro: {str(e)}', 'error')
    
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    return render_template('frota_form.html', registro=None, colaboradores=colaboradores)

@app.route('/frota/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_frota(id):
    registro = Frota.query.get_or_404(id)
    
    if request.method == 'POST':
        try:
            registro.data = datetime.strptime(request.form['data'], '%Y-%m-%d').date()
            registro.veiculo = request.form['veiculo']
            registro.motorista_id = request.form['motorista_id']
            registro.hora_saida = datetime.strptime(request.form['hora_saida'], '%H:%M').time() if request.form.get('hora_saida') else None
            registro.hora_retorno = datetime.strptime(request.form['hora_retorno'], '%H:%M').time() if request.form.get('hora_retorno') else None
            registro.km_inicial = float(request.form['km_inicial']) if request.form.get('km_inicial') else None
            registro.km_final = float(request.form['km_final']) if request.form.get('km_final') else None
            registro.observacao = request.form.get('observacao', '')
            
            # Reavalia status
            if not verificar_ponto_motorista(registro.motorista_id, registro.data):
                registro.status = 'extraordinaria'
                # Gera desconto se ainda não existe
                gerar_desconto_automatico(registro)
            else:
                registro.status = 'conforme'
            
            db.session.commit()
            flash('Registro atualizado com sucesso!', 'success')
            return redirect(url_for('frota'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar registro: {str(e)}', 'error')
    
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    return render_template('frota_form.html', registro=registro, colaboradores=colaboradores)

@app.route('/frota/excluir/<int:id>')
@login_required
def excluir_frota(id):
    registro = Frota.query.get_or_404(id)
    
    try:
        # Remove descontos automáticos associados
        Desconto.query.filter_by(frota_id=id, automatico=True).delete()
        
        db.session.delete(registro)
        db.session.commit()
        flash('Registro excluído com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir registro: {str(e)}', 'error')
    
    return redirect(url_for('frota'))

# Rotas de Descontos
@app.route('/descontos', methods=['GET'])
@login_required
def descontos():
    query = Desconto.query.order_by(Desconto.data.desc())
    colaboradores = Colaborador.query.all()
    
    # Filtros
    colaborador_id = request.args.get('colaborador_id')
    status = request.args.get('status')
    data_inicio = request.args.get('data_inicio')
    data_fim = request.args.get('data_fim')

    if colaborador_id and colaborador_id != 'all':
        query = query.filter_by(colaborador_id=int(colaborador_id))
    
    if status and status != 'all':
        query = query.filter_by(status=status)

    if data_inicio:
        query = query.filter(Desconto.data >= datetime.strptime(data_inicio, '%Y-%m-%d').date())

    if data_fim:
        query = query.filter(Desconto.data <= datetime.strptime(data_fim, '%Y-%m-%d').date())
    
    descontos = query.all()
    
    return render_template('descontos.html', descontos=descontos, colaboradores=colaboradores,
                           selected_colaborador_id=colaborador_id, selected_status=status,
                           selected_data_inicio=data_inicio, selected_data_fim=data_fim)

@app.route('/desconto/novo', methods=['GET', 'POST'])
@login_required
def novo_desconto():
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    
    if request.method == 'POST':
        try:
            desconto = Desconto(
                colaborador_id=request.form['colaborador_id'],
                data=datetime.strptime(request.form['data'], '%Y-%m-%d').date(),
                motivo=request.form['motivo'],
                valor=float(request.form['valor']),
                status=request.form['status'],
                automatico=False
            )
            
            db.session.add(desconto)
            db.session.commit()
            flash('Desconto criado com sucesso!', 'success')
            return redirect(url_for('descontos'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao criar desconto: {str(e)}', 'error')

    return render_template('desconto_form.html', desconto=None, colaboradores=colaboradores)

@app.route('/desconto/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_desconto(id):
    desconto = Desconto.query.get_or_404(id)
    colaboradores = Colaborador.query.filter_by(ativo=True).all()
    
    if request.method == 'POST':
        try:
            desconto.colaborador_id = request.form['colaborador_id']
            desconto.data = datetime.strptime(request.form['data'], '%Y-%m-%d').date()
            desconto.motivo = request.form['motivo']
            desconto.valor = float(request.form['valor'])
            desconto.status = request.form['status']
            
            db.session.commit()
            flash('Desconto atualizado com sucesso!', 'success')
            return redirect(url_for('descontos'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar desconto: {str(e)}', 'error')
            
    return render_template('desconto_form.html', desconto=desconto, colaboradores=colaboradores)

@app.route('/desconto/excluir/<int:id>')
@login_required
def excluir_desconto(id):
    desconto = Desconto.query.get_or_404(id)
    
    try:
        db.session.delete(desconto)
        db.session.commit()
        flash('Desconto excluído com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir desconto: {str(e)}', 'error')
    
    return redirect(url_for('descontos'))
    
@app.route('/desconto/aprovar/<int:id>')
@login_required
def aprovar_desconto(id):
    desconto = Desconto.query.get_or_404(id)
    desconto.status = 'aprovado'
    
    try:
        db.session.commit()
        flash('Desconto aprovado!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao aprovar desconto: {str(e)}', 'error')
    
    return redirect(url_for('descontos'))

@app.route('/desconto/cancelar/<int:id>')
@login_required
def cancelar_desconto(id):
    desconto = Desconto.query.get_or_404(id)
    desconto.status = 'cancelado'
    
    try:
        db.session.commit()
        flash('Desconto cancelado!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao cancelar desconto: {str(e)}', 'error')
    
    return redirect(url_for('descontos'))

# Rotas de Gerenciamento de Usuários (Apenas para Admin)
@app.route('/usuarios')
@login_required
def usuarios():
    if not is_admin():
        flash('Acesso negado. Apenas administradores podem gerenciar usuários.', 'error')
        return redirect(url_for('index'))
    usuarios = Usuario.query.all()
    return render_template('usuarios.html', usuarios=usuarios)

@app.route('/usuario/novo', methods=['GET', 'POST'])
@login_required
def novo_usuario():
    if not is_admin():
        flash('Acesso negado.', 'error')
        return redirect(url_for('index'))

    if request.method == 'POST':
        try:
            username = request.form['username']
            password = request.form['password']
            nome = request.form['nome']
            email = request.form['email']
            ativo = 'ativo' in request.form
            
            password_hash = generate_password_hash(password)
            
            novo_usuario = Usuario(
                username=username,
                password_hash=password_hash,
                nome=nome,
                email=email,
                ativo=ativo
            )
            db.session.add(novo_usuario)
            db.session.commit()
            flash('Usuário criado com sucesso!', 'success')
            return redirect(url_for('usuarios'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao criar usuário: {str(e)}', 'error')
            
    return render_template('usuario_form.html', usuario=None)

@app.route('/usuario/editar/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_usuario(id):
    if not is_admin():
        flash('Acesso negado.', 'error')
        return redirect(url_for('index'))

    usuario = Usuario.query.get_or_404(id)
    
    if request.method == 'POST':
        try:
            usuario.username = request.form['username']
            usuario.nome = request.form['nome']
            usuario.email = request.form['email']
            usuario.ativo = 'ativo' in request.form
            
            nova_senha = request.form.get('password')
            if nova_senha:
                usuario.password_hash = generate_password_hash(nova_senha)
            
            db.session.commit()
            flash('Usuário atualizado com sucesso!', 'success')
            return redirect(url_for('usuarios'))
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar usuário: {str(e)}', 'error')

    return render_template('usuario_form.html', usuario=usuario)

@app.route('/usuario/excluir/<int:id>')
@login_required
def excluir_usuario(id):
    if not is_admin():
        flash('Acesso negado.', 'error')
        return redirect(url_for('index'))

    usuario = Usuario.query.get_or_404(id)
    
    if usuario.username == 'admin':
        flash('Não é possível excluir o usuário administrador principal.', 'error')
        return redirect(url_for('usuarios'))
    
    try:
        db.session.delete(usuario)
        db.session.commit()
        flash('Usuário excluído com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir usuário: {str(e)}', 'error')
    
    return redirect(url_for('usuarios'))

# Exportar dados
@app.route('/exportar/<tipo>')
@login_required
def exportar(tipo):
    try:
        if tipo == 'colaboradores':
            data = Colaborador.query.all()
            df = pd.DataFrame([{
                'ID': c.id,
                'Nome': c.nome,
                'Matrícula': c.matricula,
                'CPF': c.cpf,
                'Telefone': c.telefone,
                'Email': c.email,
                'Veículo': c.veiculo_vinculado,
                'Ativo': 'Sim' if c.ativo else 'Não'
            } for c in data])
            filename = 'colaboradores.xlsx'
            
        elif tipo == 'pontos':
            data = Ponto.query.all()
            df = pd.DataFrame([{
                'ID': p.id,
                'Colaborador': p.colaborador.nome if p.colaborador else '',
                'Data/Hora': p.data_hora.strftime('%d/%m/%Y %H:%M'),
                'Tipo': p.tipo,
                'Extraordinário': 'Sim' if p.extraordinario else 'Não',
                'Observação': p.observacao
            } for p in data])
            filename = 'pontos.xlsx'
            
        elif tipo == 'frota':
            data = Frota.query.all()
            df = pd.DataFrame([{
                'ID': f.id,
                'Data': f.data.strftime('%d/%m/%Y'),
                'Veículo': f.veiculo,
                'Motorista': f.motorista_obj.nome if f.motorista_obj else '',
                'Hora Saída': f.hora_saida.strftime('%H:%M') if f.hora_saida else '',
                'Hora Retorno': f.hora_retorno.strftime('%H:%M') if f.hora_retorno else '',
                'KM Inicial': f.km_inicial,
                'KM Final': f.km_final,
                'KM Rodado': (f.km_final - f.km_inicial) if f.km_final and f.km_inicial else 0,
                'Status': f.status,
                'Observação': f.observacao
            } for f in data])
            filename = 'frota.xlsx'
            
        elif tipo == 'descontos':
            data = Desconto.query.all()
            df = pd.DataFrame([{
                'ID': d.id,
                'Colaborador': d.colaborador.nome if d.colaborador else '',
                'Data': d.data.strftime('%d/%m/%Y'),
                'Motivo': d.motivo,
                'Valor': f'R$ {d.valor:.2f}',
                'Status': d.status,
                'Automático': 'Sim' if d.automatico else 'Não'
            } for d in data])
            filename = 'descontos.xlsx'
            
        else:
            flash('Tipo de exportação inválido!', 'error')
            return redirect(url_for('index'))
        
        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados', index=False)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        flash(f'Erro ao exportar dados: {str(e)}', 'error')
        return redirect(url_for('index'))

# Inicialização do banco de dados e criação de usuário admin
@app.before_request
def create_tables():
    db.create_all()
    
    # Criar usuário admin se não existir
    if not Usuario.query.filter_by(username='admin').first():
        admin = Usuario(
            username='admin',
            password_hash=generate_password_hash('admin123'),
            nome='Administrador',
            email='admin@sistema.com',
            ativo=True
        )
        db.session.add(admin)
        db.session.commit()
        print("Usuário admin criado: username='admin', senha='admin123'")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # Criar usuário admin se não existir
        if not Usuario.query.filter_by(username='admin').first():
            admin = Usuario(
                username='admin',
                password_hash=generate_password_hash('admin123'),
                nome='Administrador',
                email='admin@sistema.com',
                ativo=True
            )
            db.session.add(admin)
            db.session.commit()
            print("Usuário admin criado: username='admin', senha='admin123'")
    
    app.run(debug=True, host='0.0.0.0', port=5007)