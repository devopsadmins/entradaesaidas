import os
import pandas
import pdfplumber
import utils


class saidas:
    _dir = None
    DENSITY = 3

    def __init__(self, dir):
        self._dir = os.getcwd() + dir

    def ler(self):
        util = utils.utils()
        arquivos = util.recursiveReadDir(self._dir)
        for arquivo in arquivos:
            if 'SAIDA' in arquivo:
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
                            if nLinha == 3:
                                cnpj = linha[73:91]
                            if nLinha == 4:
                                data_ini = linha[85:95]
                                data_fin = linha[98:108]

                            if nLinha < 13:
                                nLinha += 1
                                continue
                            if nLinha >= 13:
                                registro = self.extract_value(linha, nRegistro, registros, pagina, cnpj, data_ini,
                                                              data_fin)
                                registros.append(registro)
                                nLinha += 1
                                nRegistro += 1
        pandas.DataFrame(registros).to_excel(caminho, index=False)

    def extract_value(self, line, nRegistro, registros, pagina, cnpj, data_ini, data_fin):
        return {
            # "data_entrada": line[:10].strip() if line[:10].strip() != '' else
            # registros[nRegistro - 1]['data_entrada'],
            "especie": line[:5].strip() if line[12:20].strip() != '' else registros[nRegistro - 1][
                'especie'],
            "serie": line[9:11].strip(),
            'numero': line[12:20].strip() if line[25:35].strip() != '' else registros[nRegistro - 1][
                'numero'],
            "dia": line[25:27].strip(),
            "uf": line[28:30].strip(),
            "valor_contabil": line[31:47].strip().replace('.', ''),
            "codificacao_contabil": line[47:60].strip().replace('.', ''),
            "codificacao_fiscal": line[47:60].strip().replace('.', ''),
            "icms-ipi": line[63:67].strip(),
            "base_calculo": line[67:84].strip().replace('.', ''),
            "aliq": line[84:90].strip(),
            "imposto_debitado": line[90:105].strip().replace('.', ''),
            "nao-tributada": line[105:122].strip().replace('.', ''),
            "outras": line[122:137].strip().replace('.', ''),
            "obs": line[137:].strip(),
            "pagina": pagina.page_number,
            "cnjp": cnpj,
            "data_ini": data_ini,
            "data_fin": data_fin
        }

