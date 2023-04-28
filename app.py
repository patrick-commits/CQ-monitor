from flask import Flask, render_template, request, redirect, session
import msal
import requests
import uuid

flask_app = Flask(__name__)
flask_app.secret_key = 'your_secret_key'

# Replace these placeholders with your own values
client_id = 'your client_ID'
client_secret = 'your Client_Secret'
tenant_id = 'tenant_ID'
team_id = 'team_ID'
redirect_uri = 'http://localhost:5000/callback'

authority = f'https://login.microsoftonline.com/{tenant_id}'
scope = ['https://graph.microsoft.com/Presence.Read']

msal_app = msal.ConfidentialClientApplication(
    client_id, authority=authority, client_credential=client_secret
)

@flask_app.route('/')
def index():
    if 'user' in session:
        presence_data = get_presence_data(session['access_token'])
        return render_template('index.html', presence_data=presence_data)
    else:
        auth_url = msal_app.get_authorization_request_url(scope, redirect_uri=redirect_uri, state=str(uuid.uuid4()))
        return redirect(auth_url)

@flask_app.route('/callback')
def callback():
    code = request.args.get('code')
    result = msal_app.acquire_token_by_authorization_code(code, scope, redirect_uri=redirect_uri)
    if 'access_token' in result:
        session['user'] = result['id_token_claims']
        session['access_token'] = result['access_token']
    return redirect('/')

def get_presence_data(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    team_members_url = f'https://graph.microsoft.com/v1.0/teams/{team_id}/members'
    team_members_r = requests.get(team_members_url, headers=headers)
    team_members_json = team_members_r.json()
    team_members = team_members_json['value']

    presence_data = {}
    for member in team_members:
        user_id = member['userId']
        user_display_name = member['displayName']
        presence_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/presence'
        presence_r = requests.get(presence_url, headers=headers)
        presence_json = presence_r.json()
        #print(presence_r.json())
        presence_data[user_display_name] = presence_json
        print(presence_data)

    return presence_data

if __name__ == '__main__':
    flask_app.run(debug=True)