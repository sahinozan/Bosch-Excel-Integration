import numpy as np
import pandas as pd
import datetime

file = pd.read_excel('KW47_V00.xlsx')
file.head()

indices = file.iloc[:, [0, 7, 8, 11]].reset_index()
days = file.iloc[:, 12: 33].reset_index()
combined = pd.concat([indices, days], axis=1).iloc[2:, :]

pipes = pd.read_excel('Cihazlar - Borular.xlsx')

# pipes.head()
# combined.head()

combined = combined[combined.iloc[:, 1].notna()]

combined = combined[combined.iloc[:, 2].notna()]
combined.iloc[:, 2] = combined.iloc[:, 2].astype("str")
combined = combined[combined.iloc[:, 2].apply(str.isnumeric)]

combined = combined[combined.iloc[:, 6].apply(lambda x: (type(x) != datetime.datetime) and (type(x) != str))]
combined = combined[combined.iloc[:, 3].notna()]

combined.drop('index', axis=1, inplace=True)
combined.reset_index(drop=True, inplace=True)


initial_indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Toplam Adet"]


# In[12]:


week_days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
shifts = ["1", "2", "3"]


# In[13]:


final_indices = [" ".join([i, j]) for i in week_days for j in shifts]
initial_indices.extend(final_indices)


# In[14]:


combined.columns = initial_indices 
combined.set_index("Hat")


# In[15]:

combined["Cihaz TTNr"] = combined["Cihaz TTNr"].astype(str)
pipes["Cihaz"] = pipes["Cihaz"].astype(str)

combined = combined.merge(pipes, left_on="Cihaz TTNr", right_on="Cihaz", how="inner")
combined.drop("Cihaz", axis=1, inplace=True)
combined.insert(3, 'Boru', combined.pop('Boru'))


# In[16]:


combined.copy()


# In[17]:


for i in range(combined.shape[0]):
    combined.loc[i, "Pazartesi 1": "Pazar 3"] *= combined.loc[i, "Miktar"]

combined.to_excel("deneme.xlsx")


# In[18]:
    
yeniexcel = pd.read_excel('deneme.xlsx')

for bulunmayancihaz in file["MOE1 Üretim Sıralaması"]:
    if bulunmayancihaz not in yeniexcel["Cihaz TTNr"] :
        print(bulunmayancihaz," Bulunmuyor")
    
