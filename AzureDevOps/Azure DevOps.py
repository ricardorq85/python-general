#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#set HTTP_PROXY='http://:<psw>@proxy.epm.com.co:8080'
#set HTTP_PROXYS="http://:<psw>@proxy.epm.com.co:8080"


# In[8]:


from azure.devops.connection import Connection
from msrest.authentication import BasicAuthentication
import pprint

# Fill in with your personal access token and org URL
personal_access_token = ''
organization_url = ''

# Create a connection to the org
credentials = BasicAuthentication('', personal_access_token)
connection = Connection(base_url=organization_url, creds=credentials)


# In[9]:


def getMyProjects():

    # Get a client (the "core" client provides access to projects, teams, etc)
    core_client = connection.clients.get_core_client()
    # Get the first page of projects
    get_projects_response = core_client.get_projects()
    index = 0
    projectsList = []
    while get_projects_response is not None:
        for project in get_projects_response.value:
            #print(project)
            #pprint.pprint("[" + str(index) + "] " + project.name)
            projectsList.append(project.name)
            index += 1
        if get_projects_response.continuation_token is not None and get_projects_response.continuation_token != "":
            # Get the next page of projects
            get_projects_response = core_client.get_projects(continuation_token=get_projects_response.continuation_token)
        else:
            # All projects have been retrieved
            get_projects_response = None
    return projectsList


# In[10]:


def getReleases(project):
    release_client = connection.clients.get_release_client()
    
    releases = []
    get_releases_response = release_client.get_releases(project=project, status_filter='active')
    for release in get_releases_response.value:
        #print(release)
        releases.append(release.id)

    return releases


# In[11]:


def getManualInterventions(project_name, release):
    release_client = connection.clients.get_release_client()
    
    manual_interventions = []
    get_mi_response = release_client.get_manual_interventions(project=project_name, release_id=release)
    #print(get_mi_response)
    for mi in get_mi_response:        
        #print(mi)
        if (mi.status == 'pending'):
            manual_interventions.append("Proyecto " + project_name + "\tRelease " + str(release) + "\t " + organization_url + "/" + project_name + "/_release")
            #manual_interventions.append(mi.url)

    return manual_interventions


# In[12]:


import pymsteams

url = "https://outlook.office.com/webhook/7d8ad38b-e7be-4d10-b6c1-62b02fffed49@bf1ce8b5-5d39-4bc5-ad6e-07b3e4d7d67a/IncomingWebhook/f9ccebb148fa468dad3d08777c95b174/c4f7b00b-c27c-4a83-8a92-405ebd0f0283"
# You must create the connectorcard object with the Microsoft Webhook URL
myTeamsMessage = pymsteams.connectorcard(url)



# In[13]:


def addTeamMessage(m):
    # Add text to the message.
    myTeamsMessage.text(m)

    # send the message.
    myTeamsMessage.send()


# In[16]:


myProjects = getMyProjects()
print(myProjects)
for p in myProjects:
    releases = getReleases(p)
    for r in releases:
        manual_interventions = getManualInterventions(p, r)
        for m in manual_interventions:
            print(m)
        #    addTeamMessage(m)


# In[ ]:





