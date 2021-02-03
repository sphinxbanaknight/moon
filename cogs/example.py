import discord
import random
import os
import json
import gspread
import pprint
#import models
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
import time

from discord.ext import commands, tasks

class Example(commands.Cog):
    def __init__(self, client):
        self.client = client

    @commands.command()
    async def ping(self, ctx):
        await ctx.send(f'Pong! Latency: {round(self.client.latency * 1000)} ms')


    @commands.command(aliases = ['8ball', 'test'])
    async def _8ball(self, ctx, *, question):
        responses = ['It is certain.',
                     'It is decidedly so.',
                     'Without a doubt.',
                     'You may rely on it.',
                     'As I see it, yes.',
                     'Mostlikely.',
                     'Outlook good.',
                     'Yes.',
                     'Signs point to yes.',
                     'Reply hazy, try again.',
                     'Ask again later.',
                     'Better not tell you now.',
                     'Cannot predict now.',
                     'Concentrate, and ask again.',
                     "Don't count on it.",
                     'My reply is no.',
                     'My sources say no.',
                     'Outlook not so good.',
                     'Very doubtful.'
                    ]
        await ctx.send(f'Question: {question}\nAnswer: {random.choice(responses)}')

#    @client.command()
#    async def purge(ctx, amount=5):
#        await ctx.channel.purge(limit=amount)

def setup(client):
    client.add_cog(Example(client))