import pandas as pd
import pandapower as pp
net = pp.create_empty_network()

#For buses
df = pd.read_excel("Gendata.xlsx", sheet_name = "Bus", index_col = 0)
for idx in df.index:
    pp.create_bus(net, vn_kv = df.at[idx, "vn_kv"])
print(net.bus)
    
#for loads 
df = pd.read_excel("Gendata.xlsx", sheet_name = "Load", index_col = 0)
for idx in df.index:
    pp.create_load(net, bus = df.at[idx, "bus"], p_mw = df.at[idx, "p"])
print(net.load)

#for ext grid
df = pd.read_excel("Gendata.xlsx", sheet_name = "Gen slack", index_col = 0)
for idx in df.index:
    pp.create_ext_grid(net, bus = df.at[idx, "BUS"], p_pu = df.at[idx, "vm_pu"], va_degree=df.at[idx, "va_degree"])
print(net.ext_grid)

#for line
df = pd.read_excel("Gendata.xlsx", sheet_name = "Lines", index_col = 0)
for idx in df.index:
    pp.create_line_from_parameters(net, *df.loc[idx, :])

for transformer grid
df = pd.read_excel("Gendata.xlsx", sheet_name = "Trans", index_col = 0)
for idx in df.index:
    pp.create_transformer(net, hv_bus=df.at[idx, "H_v_bus"], lv_bus=df.at[idx, "L_v_bus"], 
                          std_type=df.at[idx, "std_type"])
print(net.transformer)
