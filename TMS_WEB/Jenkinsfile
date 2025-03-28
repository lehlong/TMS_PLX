pipeline {
    agent any
    stages {
        stage('Clone') {
            steps {
                git branch: 'staging', credentialsId: 'cf4f9fa4-db7f-4041-8d29-9086e359f55b', url: 'https://github.com/thinhgakon/WOODCHIP_DMS_WEB.git'
			}
        }
		stage('Build-Deploy') {
            steps {
                bat 'C:\\Batch-Scripts\\WOODCHIP_DMS_WEB.bat'
	              }
        }
       stage('Slack') {
         steps {
               slackSend channel: '#build', message: 'Build WOODCHIP-DMS-DATA-WEB SUCCESS!', tokenCredentialId: 'slack-bot-token'
            }
        }
	}
}	
