import os
import pandas
import pdfplumber
import utils


class entradas:
    _dir = None
    DENSITY = 3

    def __init__(self, dir):
        self._dir = os.getcwd() + dir

    def ler(self):
        util = utils.utils()
        arquivos = util.recursiveReadDir(self._dir)
        for arquivo in arquivos:
            if 'ENTRADA' in arquivo:
                # dfs = tabula.read_pdf(arquivo, pages=1)
                caminho = 'saidas-' + str(arquivo).replace(self._dir, '').replace('/', '').replace('.PDF',
                                                                                                    '') + '.csv';
                print("Arquivo a processar: " + arquivo)
                registros = []
                with pdfplumber.open(arquivo) as file:
                    for pagina in file.pages:
                        print(pagina)
                        text = pagina.extract_text(layout=False, x_density=self.DENSITY).split('\n')
                        nLinha = 0
                        nRegistro = 0
                        for linha in text:
                            if nLinha == 4:
                                cnpj = linha[59:77]
                            if nLinha == 6:
                                data_ini = linha[72:82]
                                data_fin = linha[85:95]

                            if nLinha < 14:
                                nLinha += 1
                                continue
                            if nLinha >= 14:
                                registro = self.extract_value(linha, nRegistro, registros, pagina, cnpj, data_ini,
                                                              data_fin)
                                registros.append(registro)
                                nLinha += 1
                                nRegistro += 1
        pandas.DataFrame(registros).to_excel(caminho, index=False)

    def extract_value(self, line, nRegistro, registros, pagina, cnpj, data_ini, data_fin):
        return {
            "data_entrada": line[:10].strip() if line[:10].strip() != '' else
            registros[nRegistro - 1]['data_entrada'],
            "especie": line[12:20].strip() if line[12:20].strip() != '' else registros[nRegistro - 1][
                'especie'],
            "serie": line[20:25].strip(),
            'numero': line[25:35].strip() if line[25:35].strip() != '' else registros[nRegistro - 1][
                'numero'],
            "data_documento": line[38:48].strip(),
            "codigo_emitente": line[49:64].strip(),
            "uf_orig": line[64:66].strip(),
            "valor_contabil": line[66:86].strip().replace('.', ''),
            "codificacao_contabil": line[85:95].strip().replace('.', ''),
            "codificacao_fiscal": line[85:95].strip().replace('.', ''),
            "icms-ipi": line[95:102].strip(),
            "cod": line[102:107].strip(),
            "base_calculo": line[107:125].strip().replace('.', ''),
            "aliq": line[125:132].strip(),
            "imposto_creditado": line[132:145].strip().replace('.', ''),
            "obs": line[147:].strip(),
            "pagina": pagina,
            "cnjp": cnpj,
            "data_ini": data_ini,
            "data_fin": data_fin
        }
