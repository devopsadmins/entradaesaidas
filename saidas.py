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
        novosregistros = []
        for arquivo in arquivos:
            if 'SAIDA' in arquivo:
                # dfs = tabula.read_pdf(arquivo, pages=1)
                caminho = 'saidas-.xlsx';
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
                                registro = linha
                                registros.append(self.extract_value(linha, nRegistro, registros, pagina, cnpj, data_ini,
                                                                    data_fin,arquivo))

                                nLinha += 1
                                nRegistro += 1

                        for k, registro in enumerate(registros):
                            novosregistros.append(
                                self.extract_value_array(registro, k, registros))
        with pandas.ExcelWriter(caminho) as writer:
            pandas.DataFrame(novosregistros).to_excel(writer, sheet_name='Dados', index=False)

    def extract_value(self, line, nRegistro, registros, pagina, cnpj, data_ini, data_fin,arquivo):
        return {
            "especie": line[:5].strip(),
            "serie": line[9:11].strip(),
            'numero': line[12:20].strip(),
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
            "cnpj": cnpj,
            "data_ini": data_ini,
            "data_fin": data_fin,
            "arquivo":arquivo
        }

    def extract_value_array(self, registro, nRegistro, registros):
        return {
            "especie": registro['especie'] if registro['especie'] != '' else registros[nRegistro - 1]['especie'],
            "serie": registro['serie'] if registro['serie'] != '' else registros[nRegistro - 1]['serie'],
            'numero': registro['numero'] if registro['numero'] != '' else registros[nRegistro - 1]['numero'],
            "dia": registro['dia'],
            "uf": registro['uf'],
            "valor_contabil": registro['valor_contabil'] if registro['valor_contabil'] != '' else 0,
            "codificacao_contabil": registro['codificacao_contabil'],
            "codificacao_fiscal": registro['codificacao_fiscal'],
            "icms-ipi": registro['icms-ipi'] if registro['icms-ipi'] != '' else 0,
            "base_calculo": registro['base_calculo'],
            "aliq": registro['aliq'],
            "imposto_debitado": registro['imposto_debitado'],
            "nao-tributada": registro['nao-tributada'],
            "outras": registro['outras'],
            "obs": registro['obs'],
            "pagina": registro['pagina'],
            "cnpj": registro['cnpj'],
            "data_ini": registro['data_ini'],
            "data_fin": registro['data_fin'],
            "arquivo": registro['arquivo']
        }
