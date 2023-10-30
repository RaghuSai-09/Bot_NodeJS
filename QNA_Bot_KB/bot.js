const knowledgeBase = require('./knowledge.json');
const excel = require('exceljs');
const {ActivityHandler, MessageFactory} = require('botbuilder');

class MyBot extends ActivityHandler {
    constructor(conversationState, userState) {
        super();
        this.conversationState = conversationState;
        this.userState = userState;
        this.conversationHistory = [];

        this.onMembersAdded(async (context, next) => {
            try{
                const membersAdded = context.activity.membersAdded;
                const welcomeText = 'Hello! I am Ariya, your personal health assistant. I can help you with your health concerns. How can I help you today?';
                for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                    if (membersAdded[cnt].id !== context.activity.recipient.id) {
                        await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                    }
                }
                // By calling next() you ensure that the next BotHandler is run.
                await next();
            } catch (err) {
                console.log(err);
            }
        });
    

        // Defining the main message handler.
        this.onMessage(async (context, next) => {
            // Getting the state properties from the turn context.
            const conversationData = await this.conversationState.get(context, {});
            const userData = await this.userState.get(context, {});
            const userMessage = context.activity.text;
            const response = this.matchQuestion(userMessage);

            // Storing the conversation history.
            this.conversationHistory.push({
                question: userMessage,
                answer: response
            });

            if (userMessage ==='ariya export my chart') {
                this.exportChatTranscriptToExcel(context);
                this.conversationHistory = [];
                await next();
                return;
            }

            await context.sendActivity(response || 'Sorry, I do not understand. Please try again.');
            await next();
        });
    }

    matchQuestion(question) {
        if (question === 'ariya export my chart'){
            return 'Your chat transcript has been exported successfully.';
        }
        else{
            const answer = knowledgeBase[question];
            return answer || null;
        }
    }
    exportChatTranscriptToExcel(context) {
        if (this.conversationHistory.length === 0) {
            context.sendActivity('No conversation to export.');
            return;
        }

        const workbook = new excel.Workbook();
        const worksheet = workbook.addWorksheet('Chat Transcript');
        worksheet.columns = [
            { header: 'Question', key: 'question', width: 50 },
            { header: 'Answer', key: 'answer', width: 50 },
        ];

        for (const entry of this.conversationHistory) {
            worksheet.addRow(entry);
        }

        workbook.xlsx.writeFile('chat_transcript.xlsx')
            .then(() => {
                console.log('Chat transcript exported successfully.');
            })
            .catch((error) => {
                context.sendActivity('Error exporting chat transcript.');
                console.error('Error exporting chat transcript:', error);
            });
    }
}

module.exports.MyBot = MyBot;