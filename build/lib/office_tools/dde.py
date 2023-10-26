from typing import Iterable

import win32com.client

from ..dflt import FIELDS


def get_all_fieldnames(conv, table, delim=';'):
    fields = conv.Request(f"[GetFieldNames({table}, {delim})]")
    field_list = fields.split(';')
    return field_list


# def get_all_field_values(conv, table, name):
#     fields = get_all_fieldnames(conv, table)
#     # field_num = conv.Request('[GetFieldCount(Hire)]')
#     # values = conv.Request(f"[GetFieldValues({table}, {fields}, {delim})]")
#     # values = conv.Request(f"[GetFields({table}, {name}, {fields}, )]")
#     ans = conv.Request(f"[ViewFilter(1, F,,Name, Equal to, {name},)]")
#     values = conv.Request(f"[ViewFields({table}, {name}, 2, )]")
#     return values
#
#
# def get_data_df(conv, fields: Iterable[str]) -> pd.DataFrame:
#     rows_list = []
#     for field in fields:
#         try:
#             value = conv.Request(f"[ViewField(1, {field})]")
#             row_dict = {field: value}  # Create a dictionary for the single row
#             rows_list.append(row_dict)
#         except Exception as e:
#             raise ValueError(f"Error getting {field}: {e}")
#
#     df = pd.DataFrame(rows_list)  # Create DataFrame from list of dictionaries
#     return df
#

def get_a_load(conv, table, name, fields: Iterable[str], delim=';'):
    values2 = conv.Request(f'GetFields({table}, {name}, {fields}, {delim}')


def get_data(conv, fields: Iterable[str]):
    data = {}
    for field in fields:
        try:
            data[field] = conv.Request(f"[ViewField(1, {field})]")
        except Exception as e:
            raise ValueError(f"Error getting {field}: {e}")
    return data


def fields_df(conv, fields: Iterable[str]):
    data = {}
    for field in fields:
        try:
            data[field] = conv.Request(f"[ViewField(1, {field})]")
        except Exception as e:
            raise ValueError(f"Error getting {field}: {e}")
    adf = pd.DataFrame(**data)
    ...




def get_all_connected(conv, from_item: str, connection: Connector):
    connected_names = conv.Request(
        f"[GetConnectedItemNames({connection.key_table}, {from_item}, {connection.desc}, {connection.value_table}, ;)]")
    if not connected_names or connected_names == '(none)':
        raise ValueError(
            f'{connection.value_table}:  {from_item} has no connected items  "{connection.desc}" in {connection.key_table}')
    cons = connected_names.split(';')
    results = {}
    for c_name in cons:
        connected_conv = get_record(conv, connection.table, c_name)
        connected_data = get_data(connected_conv, connection.fields)
        # results.append(connected_data)
        results[c_name] = connected_data
    return results


def get_conversation_func(topic='Commence', command='ViewData', db_name='Commence'):
    try:

        # cmc_db = win32com.client.Dispatch(f"{topic}.DB")
        cmc = win32com.client.Dispatch(f"{db_name}.DB")
        conv = cmc.GetConversation(f"{topic}", f"{command}")
    except Exception as e:
        raise ValueError(f"Could not get conversation for {topic} {command}:\n{e}")
    else:
        return conv


import win32com.client


# Usage


def get_record(conv, table, name):
    get_table(conv, table)
    ans = conv.Request(f"[ViewFilter(1, F,,Name, Equal to, {name},)]")
    item_count = conv.Request("[ViewItemCount]")
    if int(item_count) != 1:
        raise ValueError(f"{item_count} entries found for {table} : {name}")
    return conv


# def record_to_df(record_conv, table , name):
#     field_names = record_conv.Request(f"[GetFieldNames({table}, ; )]")
#     field_Count = record_conv.Request('[GetFieldCount(Hire)]')
#     ...
#
#     # record_conv.Request(f"[ViewFields(1, 1, {field_names}, ;)]")
#     ...
#
#     # field_values = get_all_field_values(conv, 'Hire', name=record_name,)
#     # return field_values
#

def get_table(conv, table):
    conv.Request(f"[ViewCategory({table})]")


#
# def items_from_hire(hire_name) -> [hire_order_items]:
#     conv = get_conversation_func()
#     record = get_record(conv, 'Hire', hire_name)
#     fields = get_all_fieldnames(conv, 'Hire')
#
#     df = get_data_df(conv, fields)
#     ...
#
#     # field_values = get_all_field_values(record, 'Hire', fields)
#     data = get_data(record, fields)
#     duration = data['Weeks']
#     items1 = []
#     for i, n in data.items():
#         if i.startswith('Number ') and int(n) > 0:
#             items1.append((i[7:], n))
#     items = [hire_order_items(i[0], i[1], duration) for i in items1 if i[0].startswith('Number ') and int(i[1]) > 0]
#     return items
#     # # order_items = (i[7:], n) for i, n in data.items() if i.startswith('Number ') and int(n) > 0)
#     # dur = 1
#     # h_order_items = hire_order_items(product_name=i[:7], quantity=int(n), duration=dur) for i, n in data.items() if i.startswith('Number ') and int(n) > 0)
#     # # oi = order_items('UHF', 10, 1)
#     # return order_items


# hire_items = items_from_hire('Test - 16/08/2023 ref 31619')
# products = get_all_hire_products(PRICES_WB)
# matched_products = match_hire_products(hire_items, products)
# a_product = list(matched_products.values())[0]
# a_price = a_product.get_price(1, 1)

#
# hire_items = products_from_hire('Test - 16/08/2023 ref 31619')
# products = get_all_hire_products(PRICES_WB)

...


class DDEContext:

    def __init__(self, topic='Commence', command='GetData'):
        self.topic = topic
        self.command = command

    def __enter__(self):
        self.dde_manager = DDEManager()
        return self.dde_manager

    def __exit__(self, exc_type, exc_val, exc_tb):
        print(f"Closing {self.topic} {self.command}")
        print(f"exc_type: {exc_type}")
        print(f"exc_val: {exc_val}")
        print(f"exc_tb: {exc_tb}")
        # self.conv = None
        ...


class DDEManager:

    def __init__(self, conversation=None):
        self.check_conv()
        self.conv = self.get_conversation(topic='Commence', command='GetData')
        ...

    def check_conv(self):
        sys_conv = self.get_conversation(topic='Commence', command='System')
        res = sys_conv.Request("Status") == 'Ready'
        assert res

    def cmc_to_df(self, table, name, fields: Iterable[str], connections: Iterable[Connection] = None):
        results = {}
        conv = self.conv
        record = get_record(conv, table, name)
        data = get_data(record, fields)
        ...
        results[table] = pd.DataFrame(**data)
        if not connections:
            pass
        else:
            for connection in connections:
                connected_data = get_all_connected(conv, table, name, connection)
                connected_df = pd.DataFrame(connected_data)
                results[connection.table] = connected_df
        return results

    def get_cmc_data(self, table, name, fields: Iterable[str], connections: Iterable[Connection] = None):
        results = {}
        conv = self.conv
        record = get_record(self.conv, table, name)
        results[table] = get_data(record, fields)
        if not connections:
            return results
        for connection in connections:
            connected_data = get_all_connected(conv, table, name, connection)
            results[connection.table] = connected_data
        return results

    def get_conversation(self, topic='Commence', command='GetData', db_name='Commence'):
        try:
            cmc = win32com.client.Dispatch(f"{db_name}.DB")
            conv = cmc.GetConversation(f"{topic}", f"{command}")
        except Exception as e2:
            print(f"Error getting conversation for {topic} {command}:\n{e2}")
            ...
        else:
            return conv

    def get_record(self, table, name):
        get_table(self.conv, table)
        res = self.conv.Request(f"[ViewFilter(1, F,,Name, Equal to, {name},)]")
        item_count = res.Request("[ViewItemCount]")
        if int(item_count) != 1:
            raise ValueError(f"{item_count} entries found for {table} : {name}")
        return res

    def get_all_fields(self, table):
        fields = self.conv.Request(f"[GetFieldNames({table}, ;)]")
        field_list = fields.split(';')
        return field_list
