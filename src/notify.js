/* 
 * Classe AppScript de Notificação criada para atender os requisitos do Fabricio Poffo
 * 
 * By Rafael Vidal | rafaelalemar@gmail.com | 20/07/2016
 * 
 * Script ID: 1iy4Lhxo1po5yT6-7kJH8DlKWylaXlDW2zVpvA0SCi7fKl7-q69F0_vK3
 */

/*
 * Carrega a biblioteca Moment.
 * Script ID: 1RsFFZZvosueH3avP3mQu6CU1TCDKITyrhXlRzfvWIlg698cYLPPJTJer
 * http://momentjs.com/
 * @type Moment Library
 */
moment = Moment.load();

Notify = {
    /*
     * Colunas utilizadas da planilha de tarefas
     * @type Dictionary
     */
    c: {empresa: 0, concluido: 1, dataLimite: 2, tarefa: 3, dataConclusao: 4, observacao: 5},
    
    /* 
     * Planilha Ativa
     * @type SpreadsheetApp
     */
    sheet: SpreadsheetApp.getActiveSpreadsheet(),
    
    /*
     * Configurações de Usuário
     * @type Dictionary
     */ 
    settings: null,
    
    /**
     * Array de tarefas atrazadas
     * @type Array
     */
    lateTasks: null,
    
    /**
     * Busca pelo settings de notificação
     * @returns {dictionary}
     */
    getSettings: function () {
        if (this.settings === null) {
            var settings = this.sheet.getSheetByName("SETTINGS").getDataRange().getValues();
            var i = 0;
            this.settings = {
                nome: settings[i++][1],
                email: settings[i++][1],
                tempo: settings[i++][1],
                subject: settings[i++][1],
                message: settings[i++][1]
            };
        }
        return this.settings;
    },
    
    /**
     * Busca por todas as tarefas do mês vigente
     * @returns {Array}
     */
    getData: function () {
        var nomePasta = moment().format("MM.YYYY");
        var sheet = this.sheet.getSheetByName(nomePasta);
        if (sheet === null) {
            throw "Planilha " + nomePasta + " não encontrada!";
        }
        return sheet.getRange("A4:F").getValues();
    },
    
    /**
     * Retorna apenas as tarefas vencidas ou com 1 dia para vencer
     * @returns {Array}
     */
    getLateTasks: function () {
        if (this.lateTasks === null) {
            var range = this.getData(), lateTasks = [], tomorrow = moment().add(this.getSettings().tempo, 'days').toDate();
            for (var i in range) {
                if (!range[i][this.c.concluido].trim() && range[i][this.c.dataLimite] <= tomorrow) {
                    lateTasks.push(range[i]);
                }
            }
            this.lateTasks = lateTasks;
        }
        return this.lateTasks;
    },
    
    /**
     * Verifica se há tarefas vencidas e manda uma notificação para o responsável
     * @returns {Notify}
     */
    check: function () {
        if (this.getLateTasks().length > 0) {
            this.sendMail();
        }
        return this;
    },
    
    /**
     * Envia uma notificação para o responsável
     * @returns {Notify}
     */
    sendMail: function () {
        var s = this.getSettings();
        var body = s.message;
        var htmlBody = "<ul>";
        
        for (var i in this.lateTasks) {
            htmlBody += "<li>" + this.lateTasks[i][this.c.empresa] + " (" + this.lateTasks[i][this.c.tarefa] + ") -> DT. " + moment(this.lateTasks[i][this.c.dataLimite]).format("DD/MM/YYYY") + "</li>";
        }
        htmlBody += "</ul>";
        
        body = body.replace("{%tarefas}", htmlBody).replace("{%nome}", s.nome);
        
        MailApp.sendEmail({
            to: s.email,
            subject: s.subject,
            htmlBody: body
        });
        return this;
    }
};