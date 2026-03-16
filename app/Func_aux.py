import re

def group_transactions(transactions, default_header_transaction, pattern_inicial):
    grouped_transactions = []
    current_transaction = []

    for line in range(len(transactions)):
        if len(transactions[line]) >= default_header_transaction and re.match(pattern_inicial, transactions[line][0]):
            if current_transaction:
                grouped_transactions.append(current_transaction)
                current_transaction = []
            current_transaction.append(transactions[line])
        else:
            current_transaction.append(transactions[line])
    if current_transaction:
        grouped_transactions.append(current_transaction)

    return grouped_transactions


def filter_transactions(transactions, exclusion_words):
    filtered_transactions = [
        transaction for transaction in transactions
        if not any(word in transaction for word in exclusion_words)
    ]
    return filtered_transactions


def _is_devolution(transaction) -> bool:
    """
    Devolução PIX contém 'DEVOLUÇÃO' explícito no texto.
    Deve ser tratada como desconto (saída) ANTES de qualquer
    verificação de PIX normal, caso contrário create_pix_entrace
    a captura e classifica errado.
    """
    transaction_str = ''.join(''.join(sub) + ' ' for sub in transaction)
    return 'DEVOLUÇÃO' in transaction_str


def process_transactions(grouped_transactions, bank_provider):
    list_discounts = []

    for line in grouped_transactions:

        # ── Devolução PIX: prioridade máxima ────────────────────────────────
        if _is_devolution(line):
            data = bank_provider.create_discount(line)
            if data:
                list_discounts.append(data)
            continue

        data = bank_provider.create_transf_sicoob(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_transf_entrace(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_dep_entrace(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_pix_entrace(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_discount(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_credit_entrace(line)
        if data:
            list_discounts.append(data)
            continue

        data = bank_provider.create_ted_entrace(line)
        if data:
            list_discounts.append(data)
            continue

    return list_discounts