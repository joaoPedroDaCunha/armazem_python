from ipp import IPP

# IP da impressora
printer_ip = "192.168.1.100"  # substitua pelo IP da sua impressora

# Criar uma instância de impressão
printer = IPP(printer_ip)

# Definir o conteúdo a ser impresso
print_data = "Data de Entrada: 05/02/2025\nNF do SAL: 12345\nLote: ABC123"

# Enviar para impressão
printer.print_data(print_data)
