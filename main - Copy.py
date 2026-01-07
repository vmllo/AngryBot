import discord
from discord.ext import commands
from discord.utils import get
from dotenv import load_dotenv
from collections import Counter
import openpyxl
import xlsxwriter
import time
import array
import os

intents = discord.Intents.default()
intents.typing = False
intents.presences = False
intents.message_content = True
user_array = []
lookup_array = []
new_message = ""
numberofthumbsup = 0
bot = commands.Bot(command_prefix='!', intents=intents)
role_id = 1249201320884305983
stuff = 0

@bot.event
async def on_ready():
    with open("usertextfile.txt",'r') as file:
        for line in file.readlines():
            user_array.append(line.strip())   
            #print(line.strip()) 
    print(f'{bot.user} has connected to Discord!')

emoji1 = bot.get_emoji('👍')
flex = bot.get_emoji('<:Flex:1199972735401607208>')
ping_message = ""

    

@bot.event
async def on_message(message):
    global stuff
    stuff = message
    global new_message
    global numberofthumbsup
    global ping_message
    global lookup_array
    guild = message.guild 
    guild_get = bot.get_guild(guild.id)
    #roles = guild.roles
    #needed_role = get(guild.roles, name="Guild_League_Ping")
    if message.content.startswith('!AB'):
        new_message = message.content.replace('!AB ','')
        new_message =  new_message.lower()
        print(new_message)
        trigger_message = new_message[0:6]
        display_message = new_message[7:]
        if trigger_message == "create":
            ping_message = await message.channel.send('\n👍: BE READY FOR THE FIRST ROUND\n\n👎: YOU ARENT GOING\n\n⛺: MEANS TENTATIVE.. UNCERTAIN; SUBJECT TO FUTURE CHANGE. IF YOU CANT MAKE IT AT FIRST ROUND PICK THIS\n\n<:Flex:1199972735401607208>: MEANS YOU CAN MAKE IT TO THE FIRST ROUND BUT ARE THERE TO FILL \n\n\nGuild League' + display_message)
            await ping_message.pin()
            await ping_message.add_reaction('👍')
            await ping_message.add_reaction('👎')
            await ping_message.add_reaction('⛺')
            await ping_message.add_reaction('<:Flex:1199972735401607208>')
        if trigger_message == "delete":
            ping_message = await message.channel.send('volc protocol starting.. DELETING MESSAGES BEEP BOOP')
            await message.channel.purge(limit=None, check=lambda msg: msg.pinned)
            await ping_message.delete()
        if trigger_message == "status":
            numberofthumbsup = len(user_array) - 1
            ping_message = await message.channel.send(str(numberofthumbsup) + "/10")
        if trigger_message == "update":
            channels = guild.channels
            workbook = xlsxwriter.Workbook('example.xlsx')
            worksheet = workbook.add_worksheet()
            display_message = new_message[7:20]
            print(display_message)
            for channel in channels:
                print(channel.name)
                if(channel.name != "NetSlum"):
                    if(channel.name != "BDO Stuff"):
                        if(channel.name != "NSFW"):
                            if(channel.name != "Voice Channels"):
                                if(channel.name != "archive"):
                                    if(channel.name != "war-shit"):
                                        if(channel.name != "Welcome"):
                                            if(channel.name != "General"):
                                                if(channel.name != "class-help"):
                                                    i = 0
                                                    async for msg in channel.history(): 
                                                        if channel.name == str(display_message): 
                                                           stuff = 'A'+ str(i)    
                                                           worksheet.write(stuff, str(msg.id)) 
                                                           i = i + 1   
                                                           print(str(msg.id))
                                                           workbook.close()                                           
        if trigger_message == "lookup":
            print("stuff")
            channels = guild.channels
            #threads = await channels.threads()
            #ping_message = await message.channel.fetch_message(msg_id)
            #await message.channel.send("stuff" + message)
            for channel in channels:
                print(channel.name)
                #if channel.name == "bdo-resources":
                if(channel.name != "NetSlum"):
                    if(channel.name != "BDO Stuff"):
                        if(channel.name != "NSFW"):
                            if(channel.name != "Voice Channels"):
                                if(channel.name != "archive"):
                                    if(channel.name != "war-shit"):
                                        if(channel.name != "Welcome"):
                                            if(channel.name != "General"):
                                                if(channel.name != "class-help"):
                                                    async for msg in channel.history():
                                                        dm = new_message[7:26]                                                     
                                                        if str(msg.id) == str(dm):                                                           
                                                            await message.channel.send(msg.jump_url)
                                                            return
            

@bot.event 
async def on_reaction_add(reaction,user):
    global user_array
    global numberofthumbsup
    global stuff
    global flag
    if reaction.message.author == bot.user:
        if reaction.emoji == '👍':
            if user.name in user_array: 
                print("")
            else:        
                user_array.append(user.name)
                with open('usertextfile.txt', 'r') as f:
                    lines = list(filter(lambda x: x.strip() != '', f.readlines()))
                with open('usertextfile.txt', 'w') as f:
                    f.writelines(lines)
                f.close
                f = open("usertextfile.txt","w+")
                for i in user_array:
                    f.write(i+"\n")
                f.close
            numberofthumbsup = len(user_array) - 1
            member = reaction.message.author
            role = get(member.guild.roles, id=role_id)
            await user.add_roles(role)

@bot.event
async def on_raw_reaction_remove(reaction):
    user = await bot.fetch_user(reaction.user_id)
    global user_array
    global numberofthumbsup
    global flag
    print(str(user))
    if numberofthumbsup > 0:
        if str(user) in user_array:
            user_array.remove(str(user))
            numberofthumbsup = len(user_array) - 1
            f = open("usertextfile.txt","w+")
            for i in user_array:
                f.write("\n" + i)
            f.close
            with open('usertextfile.txt', 'r') as f:
                lines = list(filter(lambda x: x.strip() != '', f.readlines()))
            with open('usertextfile.txt', 'w') as f:
                f.writelines(lines)
            f.close
load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')

bot.run(TOKEN)