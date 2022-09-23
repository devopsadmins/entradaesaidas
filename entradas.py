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
        novosregistros = []
        for arquivo in arquivos:
            if 'ENTRADA' in arquivo:
                # dfs = tabula.read_pdf(arquivo, pages=1)
                caminho = 'entradas-.xlsx';
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
            "data_entrada": line[:10].strip(),
            "especie": line[12:20].strip(),
            "serie": line[20:25].strip(),
            'numero': line[25:35].strip(),
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
            "cnpj": cnpj,
            "data_ini": data_ini,
            "data_fin": data_fin,
            "arquivo": arquivo,
        }

    def extract_value_array(self, registro, nRegistro, registros):
        return {
            "data_entrada": registro['data_entrada'] if registro['data_entrada'] != '' else registros[nRegistro - 1][
                'data_entrada'],
            "especie": registro['especie'] if registro['especie'] != '' else registros[nRegistro - 1]['especie'],
            "serie": registro['serie'] if registro['serie'] != '' else registros[nRegistro - 1]['serie'],
            'numero': registro['numero'] if registro['numero'] != '' else registros[nRegistro - 1]['numero'],
            "data_documento": registro['data_documento'],
            "codigo_emitente": registro['codigo_emitente'],
            "uf_orig": registro['uf_orig'],
            "valor_contabil": registro['valor_contabil'] if registro['valor_contabil'] != '' else 0,
            "codificacao_contabil": registro['codificacao_contabil'],
            "codificacao_fiscal": registro['codificacao_fiscal'],
            "icms-ipi": registro['icms-ipi'] if registro['icms-ipi'] != '' else 0,
            "cod": registro['cod'],
            "base_calculo": registro['base_calculo'],
            "aliq": registro['aliq'],
            "imposto_creditado": registro['imposto_creditado'],
            "obs": registro['obs'] if nRegistro == 0 else registros[nRegistro - 1]['obs'] + ' ' + registro['obs'],
            "pagina": registro['pagina'],
            "cnpj": registro['cnpj'],
            "data_ini": registro['data_ini'],
            "data_fin": registro['data_fin'],
            "arquivo": registro['arquivo'],
        }
