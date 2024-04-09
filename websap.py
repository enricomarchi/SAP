from attr import NOTHING
from selenium import webdriver
from selenium.webdriver import chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import InvalidSelectorException
import pandas as pd
import time
import xlwings as xw
import pymysql
import pyrfc
from sapgui import auto

TIMEOUT = 240 
hostname = 'localhost'
username = 'root'
password = '932197'
database_name = 'enricoma_lavoro'

def test(df):
    connection = connessione_mysql()
    indice_colonne = list(df.index.names)
    tutte_le_colonne = indice_colonne + list(df.columns)
    placeholders = ', '.join(['%s'] * len(tutte_le_colonne))
    colonne_sql = ', '.join(tutte_le_colonne)
    query = f"INSERT INTO sal_misure ({colonne_sql}) VALUES ({placeholders})"

    try:
        with connection.cursor() as cursor:
            for row in df.itertuples():  # Include l'indice
                # Costruisci una tupla dei valori da inserire, includendo i valori dell'indice
                dati = tuple(row.Index) + row[1:]  # 'row.Index' è una tupla contenente i valori dell'indice
                cursor.execute(query, dati)
        connection.commit()
    except Exception as e:
        print(f"Errore durante l'inserimento dei dati: {e}")
    finally:
        connection.close()

class mapping_perizie:
    def __init__(self, 
                 nome_foglio: str, 
                 riga_inizio: int, 
                 riga_fine:int, 
                 col_po: str, 
                 col_vdt:str, 
                 col_descrizione_vdt:str, 
                 col_quantità:str, 
                 col_prezzo_unitario:str, 
                 col_um:str, 
                 col_descrizione_misura:str, 
                 operazione=None, 
                 parte_opera=None):
        self.nome_foglio = nome_foglio
        self.riga_inizio = riga_inizio
        self.riga_fine = riga_fine
        self.col_po = col_po
        self.col_vdt = col_vdt
        self.col_descrizione_vdt = col_descrizione_vdt
        self.col_quantità = col_quantità
        self.col_prezzo_unitario = col_prezzo_unitario
        self.col_um = col_um
        self.col_descrizione_misura = col_descrizione_misura
        self.parte_opera = parte_opera
        self.operazione = operazione

class Ex_VDT_non_trovata(Exception):
    def __init__(self, vdt, messaggio):
        super().__init__(messaggio)
        self.vdt = vdt

class Ex_xpath_vuoto(Exception):
    def __init__(self, messaggio):
        super().__init__(messaggio)

def colonna_da_nome(ws, nome):
      col_ref = ws.range(nome).address

      # Estrai la lettera della colonna dal riferimento alla cella
      col_letter = ''.join(c for c in col_ref if c.isalpha())

      # Converti la lettera della colonna nel numero corrispondente
      col_num = 0
      for char in col_letter:
            col_num = col_num * 26 + (ord(char.upper()) - ord('A')) + 1

      return col_num

def connessione_mysql():
    return pymysql.connect(host=hostname, user=username, password=password, db=database_name)

def engine_sqlalchemy():
    engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@{hostname}/{database_name}')
    return engine

def elimina_sal(id_sal):
    execute_query("DELETE FROM sal_misure WHERE id_sal = %s", args=(id_sal,)) 

def salva_sal(df):
    engine = engine_sqlalchemy()
    df.columns = df.columns.str.lower() 
    df.to_sql('sal_misure', con=engine, if_exists='append', index=True)
    engine.dispose()      

def elimina_perizia(id_perizia):
    execute_query(f"DELETE FROM perizie_misure WHERE id_perizia = %s", args=(id_perizia,)) 

def salva_perizia(df):
    for index, row in df.iterrows():
        # Aggiunta degli indici alla lista dei valori
        values_with_indices = [index[0], index[1]] + [None if (pd.isna(value)) else value for value in row]
        
        # Aggiunta degli indici alla lista dei nomi delle colonne
        columns_with_indices = ['id_perizia', 'n_riga'] + df.columns.tolist()
        
        # Creazione della query di inserimento con i parametri
        sql = f"INSERT INTO perizie_misure ({', '.join(columns_with_indices)}) VALUES ({', '.join(['%s' for _ in range(len(values_with_indices))])})"
        
        # Esecuzione della query utilizzando la funzione execute_query
        execute_query(sql, tuple(values_with_indices))

def aggiorna_riga_sal(index, row):
    sql = 'UPDATE sal_misure SET oda=%s, n_sal=%s, posizione=%s, po=%s, descrizione=%s, vdt=%s, quantità=%s, nv=%s, prezzo_nv=%s, um_nv=%s, rib=%s, data_misura=%s, edizione_tariffa=%s, inserita=%s, note=%s, edizione_tariffa_adeguamento=%s, sovrapprezzo1=%s, sovrapprezzo2=%s WHERE id_sal=%s AND n_riga=%s'
    args = (row.oda, row.n_sal, row.posizione, row.po, row.descrizione_po, row.descrizione, row.vdt, row.quantità, row.nv, row.prezzo_nv, row.um_nv, row.rib, row.data_misura, row.edizione_tariffa, row.inserita, row.note, row.edizione_tariffa_adeguamento, row.sovrapprezzo1, row.sovrapprezzo2, index[0], index[1])
    execute_query(sql, args)
    
def read_stored_proc(nome, args):
    conn = connessione_mysql()
    try:
        with conn.cursor() as cursor:
            cursor.callproc(nome, args=args)
            result = cursor.fetchall()
            columns = [col[0] for col in cursor.description]
            df = pd.DataFrame(result, columns=columns)
            return df
    finally:
        conn.close()    

def execute_stored_proc(nome, args):
    conn = connessione_mysql()
    try:
        with conn.cursor() as cursor:
            cursor.callproc(nome, args=args)
            conn.commit() 
    except Exception as e:
        print(f"Si è verificato un errore: {e}")
        conn.rollback()
    finally:
        conn.close()

def read_query(sql, args):
    conn = connessione_mysql()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, args)
            result = cursor.fetchall()
            columns = [col[0] for col in cursor.description]
            df = pd.DataFrame(result, columns=columns)            
            return df
    finally:
        conn.close()    

def execute_query(sql, args):
    conn = connessione_mysql()
    args = list(map(lambda x: x.replace("'", "\'") if isinstance(x, str) else x, args))
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, args)
            conn.commit()
    except Exception as e:  
        print(f"Errore durante l'esecuzione della query: {e}")
        conn.rollback()  
    finally:
        conn.close()    

def importa_sal_da_excel(file_excel, append=True):
    df = pd.read_excel(io=file_excel, sheet_name="VDT", header=0)
    id_sal = df['id_sal'].iloc[0]
    if append:
        temp_df = read_query("SELECT MAX(n_riga) AS max_riga FROM sal_misure WHERE id_sal=%s GROUP BY id_sal", args=(id_sal,))
        inizia_da_riga = temp_df['max_riga'].iloc[0] + 1
    else:
        elimina_sal(id_sal)
        inizia_da_riga = 1
    df.insert(1, 'n_riga', df.reset_index().index + inizia_da_riga)
    df['data_misura'] = pd.to_datetime(df['data_misura'], format='%d.%m.%Y').dt.strftime('%Y-%m-%d')
    df.set_index(['id_sal', 'n_riga'], inplace=True)
    salva_sal(df)

def importa_perizia_da_excel(file_excel: str, id_perizia: str, network:str, *mapping: mapping_perizie, append: bool = True):
    try:
        app = xw.App(add_book=False, visible=False)
        wb = app.books.open(file_excel)
        df_perizia = pd.DataFrame(columns=['id_perizia', 'operazione', 'po', 'descrizione_misura', 'vdt', 'descrizione_vdt', 'quantità', 'tipo_vdt', 'prezzo_nv', 'um_nv', 'inserita', 'sovrapprezzo1', 'sovrapprezzo2', 'foglio_excel', 'riga_excel'])

        for mappa in mapping:
            print(f'\n{mappa.nome_foglio}', end='\n', flush=True)
            ws = wb.sheets[mappa.nome_foglio]
            for riga in range(mappa.riga_inizio, mappa.riga_fine+1):
                print(f'\rRiga {riga} di {mappa.riga_fine}', end='', flush=True)
                vdt = ws.range(f'{mappa.col_vdt}{riga}').value if mappa.col_vdt else None
                descrizione_vdt = ws.range(f'{mappa.col_descrizione_vdt}{riga}').value if mappa.col_descrizione_vdt else None
                quantità = ws.range(f'{mappa.col_quantità}{riga}').value if mappa.col_quantità else None
                prezzo_unitario = ws.range(f'{mappa.col_prezzo_unitario}{riga}').value if mappa.col_prezzo_unitario else None
                um = ws.range(f'{mappa.col_um}{riga}').value if mappa.col_um else None
                descrizione_misura = ws.range(f'{mappa.col_descrizione_misura}{riga}').value if mappa.col_descrizione_misura else None
                if mappa.col_po:
                    po = ws.range(f'{mappa.col_po}{riga}').value
                elif mappa.parte_opera:
                    po = mappa.parte_opera
                else:
                    po = 'Unica'
                
                operazione = mappa.operazione if mappa.operazione else '10'
                
                riga_df = pd.Series()
                if (vdt or descrizione_vdt or quantità):
                    riga_df['id_perizia'] = id_perizia
                    riga_df['operazione'] = operazione
                    riga_df['po'] = po
                    if descrizione_misura: 
                        riga_df['descrizione_misura'] = descrizione_misura
                    if vdt:
                        riga_df['vdt'] = vdt
                    else:
                        riga_df['vdt'] = descrizione_vdt
                    if descrizione_vdt:
                        riga_df['descrizione_vdt'] = descrizione_vdt
                    if quantità:
                        riga_df['quantità'] = quantità
                    if prezzo_unitario:
                        riga_df['prezzo_nv'] = prezzo_unitario
                    if um:
                        riga_df['um_nv'] = um
                    riga_df['foglio_excel'] = mappa.nome_foglio
                    riga_df['riga_excel'] = riga
                        
                    df_perizia.loc[len(df_perizia)] = riga_df
                
        if append:
            temp_df = read_query("SELECT MAX(n_riga) AS max_riga FROM perizie_misure WHERE id_perizia=%s GROUP BY id_perizia", args=(id_perizia,))
            inizia_da_riga = temp_df['max_riga'].iloc[0] + 1
        else:
            elimina_perizia(id_perizia)
            inizia_da_riga = 1

        df_perizia.dropna(subset=['quantità'], inplace=True)
        df_perizia = df_perizia[df_perizia['quantità'] != 0]
        df_perizia.insert(1, 'n_riga', df_perizia.reset_index().index + inizia_da_riga)
        df_perizia.set_index(['id_perizia', 'n_riga'], inplace=True)
        
        salva_perizia(df_perizia)
        
        return df_perizia
    except Exception as e:
        print(str(e))
    finally:
        wb.close()
        app.quit()
        
class WebSAP:
    DEF_tempo_operazione: float = 1.5
    DEF_timeout: float = 600

    def __init__(
            self,
            utente: str,
            password: str
    ):
        """
        Lancia l'applicativo Web PS2 e si connette con il nome utente e password forniti.
        :param utente: nome utente
        :param password: password
        """
        self.utente = utente
        self.password = password
        self.driver = None
        self.tempo_operazione = WebSAP.DEF_tempo_operazione
        self.timeout = WebSAP.DEF_timeout
        self.wb = None
        self.ws = None

    def logon(
            self
    ):
        """
        Esegue il logon con l'utenza e la password fornita durante la creazione dell'oggetto
        :return:
        """
        self.driver: chrome = webdriver.Chrome()
        self.driver.maximize_window()
        self.driver.get("https://ps2.rfi.it/sap/bc/webdynpro/sap/z0029_wd_rfi_main_comp?sap-client=111&sap-language=IT#")
        self.testo(self.utente, '//input[@id="sap-user"]')  # Utente
        self.testo(self.password, '//input[@id="sap-password"]')  # Password
        self.click('//div[@id="LOGON_BUTTON"]')  # LOGON

        lista = self.driver.find_elements(By.XPATH, '//span[contains(text(), "utente è già collegato al sistema con le seguenti sessioni")]')
        if len(lista) > 0:
            self.click('//span[text()="Cont."]/../..')

        self.click('//div[@id="SYSTEM_MESSAGE_CONTINUE_BUTTON"]')

    def elemento(
            self,
            xpath: str
    ):
        """
        Ritorna l'elemento HTML con xpath indicato.
        :param xpath: xpath dell'elemento desiderato
        :return: WebElement
        """
        for i in range(1, TIMEOUT):
            try:
                elem = self.driver.find_element(By.XPATH, xpath)
                return elem
            except Exception as e:
                #print(str(e))
                print(f'\rProc. elemento: xpath = {xpath}, Tentavivo n.{i} di {TIMEOUT}', end=' ', flush=True)
                time.sleep(1)
        print('\n', end='\n')
        #elem = WebDriverWait(self.driver, self.timeout, 1).until(lambda x: x.find_element(By.XPATH, xpath))

    def lista_elementi(
            self,
            xpath: str
    ):
        """
        Ritorna una lista di elementi HTML con xpath indicato.
        :param xpath: xpath degli elementi desiderati
        :return: WebElement
        """
        lista = []
        for i in range(1, TIMEOUT):
            try:
                lista = self.driver.find_elements(By.XPATH, xpath)
                return lista
            except NoSuchElementException as e:
                #print(str(e))
                print(f'\rProc. lista_elementi: xpath = {xpath}, Tentavivo n.{i} di {TIMEOUT}', end=' ', flush=True)
                time.sleep(1)
        print('\n', end='\n')
        return lista

    def attesa_caricamento(
            self,
    ):
        """
        Attende il caricamento della pagina, fino a quando non scompare la ruota blu che gira.
        """
        time.sleep(1)
        elem: WebElement = self.elemento('//div[@id="ur-loading"]')
        while elem.value_of_css_property("visibility") == "visible":
            time.sleep(1)

    def click(
            self,
            xpath: str,
    ):
        """
        Simula un click del mouse sull'elemento puntato da xpath.
        :param xpath: xpath dell'elemento da cliccare
        :return: WebElement
        """
        if xpath:
            for i in range(1, TIMEOUT):
                try:
                    elem = self.driver.find_element(By.XPATH, xpath)
                    elem.click()
                    return elem        
                except Exception as e:
                    #print(str(e))
                    print(f'\rProc. click: xpath = {xpath}, Tentavivo n.{i} di {TIMEOUT}', end=' ', flush=True)
                    time.sleep(1)
        else:
            raise Ex_xpath_vuoto('Nessun xpath fornito')
        print('\n', end='\n')
        
    def testo(
            self,
            txt: str,
            xpath: str
    ):
        """
        Fa un click sull'elemento ed inserisce il testo specificato.
        :param txt: il testo da inserire
        :param xpath: xpath dell'elemento
        :return: WebElement
        """
        if xpath:
            for i in range(1, TIMEOUT):
                try:
                    elem = self.click(xpath)
                    time.sleep(1)
                    elem.clear()
                    time.sleep(1)
                    elem.send_keys(txt)
                    time.sleep(1)
                    return elem
                except Exception as e:
                    #print(str(e))
                    print(f'\rProc. click: xpath = {xpath}, Tentavivo n.{i} di {TIMEOUT}', end=' ', flush=True)
                    time.sleep(1)
        else:
            raise Ex_xpath_vuoto('Nessun xpath fornito')
        print('\n', end='\n')
                    
    def sal(self, id_sal, riepilogo_vdt=1, rielaborazione=0):
        oda = id_sal[:10]
        # Entra in WebSAL
        time.sleep(1)
        self.click('//div[@title="SAL"]')           # Tasto SAL
        time.sleep(1)
        self.click('//td[@ut="3"][@cc="3"]/a')      # Tasto freccia Assistente
        time.sleep(1)
        self.click('//span[text()="OK"]/../..')     # OK al popup
        if rielaborazione == 0:
            self.click('//div[text()="GESTIONE"]')      # Scheda GESTIONE
        else:
            self.click('//div[text()="RIELABORAZIONE"]') # Scheda RIELABORAZIONE
            
        self.attesa_caricamento()

        df = read_stored_proc('vdt_da_inserire', (id_sal, riepilogo_vdt))
        if len(df)>0:
            if riepilogo_vdt == 0:
                df.set_index(['id_sal', 'n_riga'], inplace=True)
            else:
                df.set_index(['id_sal'], inplace=True)
                df['descrizione'] = ''
            
            time.sleep(1)       
            self.testo(oda, '//span[text()="NUMERO ORDINE DI ACQUISTO NOTO"]/../../../following-sibling::td//input')  # ODA

            df = df.dropna(axis=0, subset=['quantità'])
            df.vdt = df.vdt.str.upper()
            df.vdt = df.vdt.fillna("")
            df.nv = df.nv.fillna("")
            df.descrizione = df.descrizione.fillna("")
            df.inserita = df.inserita.fillna("")
            posizioni = df["posizione"].unique()
            tot_vdt = len(df)
            i_vdt = 1
            for pos in posizioni:
                self.testo(str(pos), '//span[text()="Posizione Numero"]/../../../following-sibling::td//input')  # Posizione
                df_posizione = df.loc[df["posizione"] == pos, :]
                parti_opera = df_posizione["po"].unique()
                for po in parti_opera:
                    self.testo(str(po), '//span[contains(text(), "Opera")]/../../../following-sibling::td//input')   # Parte d'opera
                    df_parte_opera = df_posizione.loc[df_posizione["po"] == po, :]
                    if rielaborazione == 0:
                        self.click('//div[@title="Vai alla Gestione delle Voci di Tariffa"]')  # Gestione misurazioni
                    else:
                        self.click('//div[@title="Procedi con la gestione delle Misurazioni in Rielaborazione"]')  # Gestione misurazioni in RIELAB.
                    self.attesa_caricamento()

                    # Se compare una richiesta di modificare la versione in elaborazione clicca e prosegui
                    time.sleep(1)
                    lista = self.lista_elementi('//span[text()="Ultima versione IN ELABORAZIONE"]/..')
                    if len(lista) > 0:
                        lista[0].click()  # Checkbox "ultima versione in elaborazione"
                        self.click('//span[text()="OK"]/../..')  # Tasto OK
                        self.attesa_caricamento()

                    time.sleep(1)
                    for index, row in df_parte_opera.iterrows():
                        if riepilogo_vdt == 0:
                            n_riga = index[1]
                        else:
                            n_riga = 0
                        if row.inserita == "":
                            try:
                                print(f"VDT n.{i_vdt} di {tot_vdt}, VDT = {row.vdt}, Q = {row.quantità}, Descrizione = {row.descrizione}")
                                time.sleep(5)
                                self.attesa_caricamento()
                                self.__inserisci_vdt_sal(index, row)
                                time.sleep(1)
                                execute_stored_proc('modifica_vdt_inserita_in_sal', (id_sal, row.vdt, n_riga, 'x'))
                            except Ex_VDT_non_trovata as e:
                                execute_stored_proc('modifica_vdt_inserita_in_sal', (id_sal, row.vdt, n_riga, 'manca'))
                                print(str(e))
                            finally:
                                i_vdt += 1

                    self.click('//span[text()="Salva"]/../..')              # SALVA
                    self.click('//span[text()="Altra gestione"]/../..')     # Altra gestione    
        else:
            print(f'SAL {id_sal} non trovata nel database, oppure tutte le voci sono già state inserite')               

    def __inserisci_vdt_sal(self, index, row):
        vdt = row.vdt
        q = row.quantità
        data = row.data_misura.strftime('%d.%m.%Y')
        ribasso = row.rib
        descrizione = row.descrizione
        nuova_voce = row.nv
        ed_tariffa = row.edizione_tariffa

        # Verifica se la VDT è già stata inserita
        time.sleep(1)
        if nuova_voce:
            lista = self.lista_elementi(f'//span[contains(translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "{vdt}")]')
        else:
            lista = self.lista_elementi(f'//td[3]//span[translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")="{vdt}"]')  

        time.sleep(1)
        if len(lista) == 0:                                 # Se non è stata già inserita la inserisce
            self.click('//div[@title="Seleziona una nuova Voce di Tariffa"]')
            time.sleep(1)
            self.click('//div[text()="Contr.Catalogo"]')
            time.sleep(1)
            if nuova_voce:
                self.testo(f'*{vdt}*', '//div[@class="urPWContent"]//span[text()="Descrizione"]/../../../../td[4]//input')  # campo Descrizione
                elem = self.elemento('//div[@class="urPWContent"]//span[text()="VdT"]/../../../../td[2]//input')  # cancella il campo VDT
                elem.clear()
            else:
                self.testo(vdt, '//div[@class="urPWContent"]//span[text()="VdT"]/../../../../td[2]//input')  # campo VDT
                elem = self.elemento('//div[@class="urPWContent"]//span[text()="Descrizione"]/../../../../td[4]//input')  # cancella il campo descrizione
                elem.clear()
            time.sleep(1)
            self.click('//span[text()="Visualizza"]/../..')     # visualizza

            # ============================================ SELEZIONE EDIZIONE TARIFFA =========================================
            try:
                time.sleep(1)
                versione_vdt = self.trova_versione_vdt(index, row)
            except Ex_VDT_non_trovata:
                time.sleep(1)
                self.click('//span[text()="Chiudi"]/../..')
                raise
            time.sleep(1)
            versioni = self.lista_elementi(
                f'//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr[@rr="{versione_vdt}"]/td[@cc="0"]'
            )
            time.sleep(1)
            if len(versioni) > 0:
                versioni[0].click()   # checkbox voce
                elem = self.elemento(
                    f'//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr[@rr="{versione_vdt}"]/td[3]/span/span')
            else:  # in caso contrario preme la prima della lista, che dovrebbe essere l'unica dell'elenco
                self.click(f'//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr[@rr="1"]/td')
                elem = self.elemento(f'//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr[@rr="1"]/td[3]/span/span')

            time.sleep(1)
            # =================================================================================================================
    
            self.click('//span[text()="Inserisci"]/../..')      # inserisci

        time.sleep(1)
        if nuova_voce:
            self.click(f'//td[4]//span[contains(translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "{vdt}")]/../../../../../../../../../../../../td[2]//span[text()="Gestione Misurazioni"]/../..')  # Gestione misurazioni
        else:   
            self.click(f'//td[3]//span[translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")="{vdt}"]/../../../../../../../../../../../../td[2]//span[text()="Gestione Misurazioni"]/../..')  # Gestione misurazioni

        time.sleep(1)
        tabella_misure = '//span[text()="+/-"]/../../../../..'  # xpath della tabella misure
        tot_righe = self.lista_elementi(f'{tabella_misure}/tr')  # N.B. la riga n. 1 è l'header
        # Trova le colonne interessate nell'header
        time.sleep(1)
        header = self.lista_elementi('//span[text()="+/-"]/../../../../../tr[1]/th')
        indice_link_descrizione = 0
        indice_misura = 0
        indice_data = 0
        time.sleep(1)
        for i in range(1, len(header)):
            elem = self.elemento(f'//span[text()="+/-"]/../../../../../tr[1]/th[{i}]//span/span')
            time.sleep(1)
            if elem.text == "Nota":
                indice_link_descrizione = i
            elif elem.text == "Totale":
                indice_misura = i
            elif elem.text[0:4] == "Data":
                indice_data = i
                break
        xpath_data = ""
        xpath_link_descrizione = ""
        xpath_misura = ""

        # Cerca la prima riga vuota dove inserire la misura
        time.sleep(1)
        for i in range(2, len(tot_righe)):   # la riga n.1 è l'header
            #time.sleep(1)
            xpath_misura = f'{tabella_misure}/tr[{i}]/td[{indice_misura}]//input'
            campo_misura: WebElement = self.elemento(xpath_misura)
            xpath_data = f'{tabella_misure}/tr[{i}]/td[{indice_data}]//input'
            xpath_link_descrizione = f'{tabella_misure}/tr[{i}]/td[{indice_link_descrizione}]//a'
            if campo_misura.get_attribute('value') == "":   # quando trova la riga vuota esce dal ciclo
                break

        self.testo(data, xpath_data)
        time.sleep(1)
        self.testo(str(q) + Keys.ENTER, xpath_misura)
        time.sleep(1)
        if descrizione != "":
            self.click(xpath_link_descrizione)
            self.testo(descrizione, '//textarea')
            self.click('//span[text()="Salva"]/../..')  # Salva
        time.sleep(1)
        errore = 1
        while errore == 1:
            try:
                self.click('//span[text()="Torna alla SAL"]/../..')     # Torna alla sal
                errore = 0
            except Ex_xpath_vuoto as e:
                errore = 1
        if ribasso == "NO":
            time.sleep(2)
            if nuova_voce:
                self.click(f'//td[4]//span[contains(translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "{vdt}")]/../../../../../../../../../../../../..//span[text()="No"]/..')
            else:    
                self.click(f'//td[3]//span[translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")="{vdt}"]/../../../../../../../../../../../../..//span[text()="No"]/..')
            time.sleep(2)
            self.click('//span[text() = "Aggiorna"] /../..')
            self.attesa_caricamento()
            time.sleep(2)

    def trova_versione_vdt(self, index, row):
        nuova_voce = row.nv
        vdt = row.vdt
        ed_tariffa = row.edizione_tariffa
        tot_righe_elem = self.lista_elementi('//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr/td/table/tbody/tr')
        tot_righe = len(tot_righe_elem) - 1 # -1 perché una riga è l'header 
        if tot_righe == 0:
            raise Ex_VDT_non_trovata(vdt, f"La VDT {vdt} edizione {ed_tariffa} non è presente in web sal.")
        elif nuova_voce == "x":
            return 1
        else:
            prezzo_da_trovare = round(row.prezzo_unitario, 2)
        trovata = False
        for riga in range(1, tot_righe + 1):
            time.sleep(1)
            pr_un = self.elemento(f'//span[text()="Visualizza"]/../../../../../../../../../../../tr[2]//td//td//tr[2]//tr[@rr="{riga}"]/td[@cc="6"]')
            time.sleep(1)
            pr_un = pr_un.text
            pr_un = pr_un.replace(".", "").replace(",", ".")
            pr_un = float(pr_un)
            if pr_un == prezzo_da_trovare:
                trovata = True
                break
        if trovata == False:
            raise Ex_VDT_non_trovata(vdt, f"La VDT {vdt} edizione {ed_tariffa} non è presente in web sal.")
        
        return riga
            
    def perizia(self, network: str, file_excel: str, nome_foglio: str):
        # Entra in Web Perizie
        time.sleep(1)
        self.click('//div[@title="Perizie"]')       # Tasto perizie
        time.sleep(1)
        self.click('//td[@ut="3"][@cc="3"]/a')      # Tasto freccia Specialista
        self.attesa_caricamento()
        self.click('//span[text()="OK"]/../..')     # OK al popup
        self.click('//div[text()="GESTIONE VPR"]')      # Scheda GESTIONE
        self.attesa_caricamento()
        self.testo(network, '//span[text()="NUMERO DI NETWORK"]/../../../following-sibling::td//input')  # Network

        df = pd.read_excel(io=file_excel, sheet_name="VDT", header=0)
        df = df[(df["Inserita"] == "") | (pd.isna(df["Inserita"]))]
        df.vdt = df.vdt.str.upper()
        df.vdt = df.vdt.fillna("")
        df.descrizione = df.descrizione.fillna("")
        df = df[df["Quantità"] != 0]
        df.inserita = df.inserita.fillna("")
        df.Prezzo = df.Prezzo.fillna(0)
        operazioni = df.Operazione.unique()
        for op in operazioni:
            self.testo(str(op), '//span[text()="OPERAZIONE"]/../../../following-sibling::td//input')  # Operazione
            df_operazione = df[df["Operazione"] == op]
            parti_opera = df_operazione.PO.unique()
            for po in parti_opera:
                elem = self.testo(str(po), '//span[starts-with(text(), "PARTE")]/../../../following-sibling::td//input')   # Parte d'opera
                df_parte_opera = df_operazione[df_operazione["PO"] == po]
                elem.send_keys(Keys.RETURN)
                self.click('//span[text()="GESTIONE RISORSE"]/../..')
                self.attesa_caricamento()

                # Se compare un messaggio che un altro utente sta modificando questa VPR, seleziona di elaborare questa VPR
                time.sleep(1)
                lista = self.lista_elementi('//span[starts-with(text(), "Continuare con questa V.P.R.")]/..')
                if len(lista) > 0:
                    lista[0].click()  # Checkbox "Continuare con questa VPR"
                    self.click('//span[text()="PROCEDI"]/../..')  # Tasto PROCEDI
                    self.attesa_caricamento()

                # Se compare un messaggio che sono state rilevate versioni non salvate
                time.sleep(1)
                lista = self.lista_elementi('//span[starts-with(text(), "Ultima versione IN ELABORAZIONE")]')
                if len(lista) > 0:
                    lista[0].click()  # Checkbox "Ultima versione in elaborazione"
                    self.click('//span[text()="PROCEDI"]/../..')  # Tasto PROCEDI
                    self.attesa_caricamento()

                time.sleep(1)
                for index, row in df_parte_opera.iterrows():
                    if row.Inserita == "":
                        try:
                            self.__inserisci_vdt_perizia(index, row)
                            df.at[index, "Inserita"] = "x"
                            self.ws.cell(row=index, column=df.columns.get_loc("Inserita")+1).value = "x"
                        except InvalidSelectorException as e:
                            pass
                        finally:
                            self.wb.save(file_excel)
                            
                self.click('//span[text()="SALVA"]/../..')              # SALVA
                self.click('//span[text()="ALTRA GESTIONE"]/../..')     # Altra gestione
                self.wb.close()
                
    def __inserisci_vdt_perizia(self, index, row):
        vdt = row.VDT
        q = row.Quantità
        prezzo=row.Prezzo
        vdt_standard = row.TipoVDT == "ANAGRAFICA"
        descrizione = row.Descrizione
        UM = row.UM
        
        # Verifica se la VDT è già stata inserita
        print(vdt)
        time.sleep(1)
        if vdt_standard:
            lista = self.lista_elementi(f'//input[@value="{vdt}"]')
        else:
            lista = self.lista_elementi(f'//span[text()="{vdt}"]')

        time.sleep(1)
        if len(lista) == 0:  # Se non è stata già inserita la inserisce
            self.click('//span[text()="[SEL. VdT]"]/..')  # Sel. VDT
            time.sleep(1)
            if vdt_standard:
                self.testo(vdt, '//div[@class="urPWContent"]//span[text()="Voce di Tariffa"]/../../../td[5]//input')  # campo VDT
                self.click('//span[text()="VISUALIZZA"]/../..')  # visualizza
                time.sleep(1)
                self.click(f'//span[text()="{vdt}"]/../../../td')  # checkbox prima voce
            else:
                self.click('//div[text()="NON CODIFICATE"]')
                self.testo(vdt, '//div[@class="urPWContent"]//span[text()="DESCRIZIONE"]/../../../../../../../../../../tr[2]/td[3]//input')  # campo Descrizione
                self.testo(UM, '//div[@class="urPWContent"]//span[text()="DESCRIZIONE"]/../../../../../../../../../../tr[2]/td[4]//input')  # campo UM
            time.sleep(1)
            self.click('//span[text()="INSERISCI"]/../..')  # inserisci
        time.sleep(1)
        if vdt_standard:
            self.click(f'//input[@value="{vdt}"]/../../../td[8]/a')  # Gestione misurazioni
        else:
            # se è una VDT non codificata inserisce anche l'importo unitario
            self.testo(str(prezzo), f'//span[text()="{vdt}"]/../../../../../../../../../../../../../../../../../../tr[2]//span[text()="Prezzo Lordo :"]/../../../td[4]//input')   # campo Prezzo
            time.sleep(1)
            self.click(f'//span[text()="{vdt}"]/../../../td[8]/a')  # Gestione misurazioni

        time.sleep(1)
        tabella_misure = '//span[text()="+/-"]/../../../../..'  # xpath della tabella misure
        tot_righe = []
        try:
            tot_righe = self.lista_elementi(f'{tabella_misure}/tr')  # N.B. la riga n. 1 è l'header
        except TimeoutException as e:
            print(f"TimeoutException, xpath:{tabella_misure}/tr , def __inserisci_vdt_perizia, riga 391")
        xpath_link_descrizione = ""
        xpath_misura = ""
        # Cerca la prima riga vuota dove inserire la misura
        time.sleep(1)
        for i in range(2, len(tot_righe)):  # la riga n.1 è l'header
            xpath_misura = f'{tabella_misure}/tr[{i}]/td[5]//input'
            campo_misura: WebElement = self.elemento(xpath_misura)
            xpath_link_descrizione = f'{tabella_misure}/tr[{i}]/td[3]/div'
            if campo_misura.get_attribute('value') == "":  # quando trova la riga vuota esce dal ciclo
                break

        time.sleep(1)
        self.testo(str(q) + Keys.ENTER, xpath_misura)
        if descrizione != "":
            self.click(xpath_link_descrizione)
            self.testo(descrizione, '//textarea')
            self.click('//span[text()="[SALVA]"]/../..')  # Salva
        time.sleep(1)
        self.click('//span[text()="[TORNA ALLA VPR]"]/../..')  # Torna alla VPR

def test_sap():
    # Connessione a SAP
    conn = pyrfc.Connection(
        user='932197',
        passwd='.Gioia7777',
        ashost='in00.rfi.it',
        sysnr='02',
        client='111',
        lang='IT'  # Lingua
    )

    # Ottenere l'oggetto SAP GUI
    sapgui_auto = auto.find_sapgui()
    app = sapgui_auto.GetScriptingEngine

    # Ottenere la connessione e la sessione
    connection = app.Children(0)
    session = connection.Children(0)

    # Ora puoi interagire con la sessione, ad esempio:
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nZGEST_FND"
    session.findById("wnd[0]").sendVKey(0)  # Invia il tasto 'Enter'
