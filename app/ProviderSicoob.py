from app.ProviderBank import ProviderBank
import re

class ProviderSicoob(ProviderBank):
    def __init__(self, pattern_inicial=None, group_words_clear=None):
        super().__init__()
        self.pattern_inicial = pattern_inicial
        self.group_words_clear = group_words_clear

    def create_discount(self, transaction):
        item_for_convert_to_discount = []
        transaction_type = ' '.join(transaction[0][1:-1])
        words_payment = ['DEB', 'DÉB', 'VISA', 'COMP']

        if self._verify_discount(transaction):
            payment_form = self._verify_payment_form(transaction[0], words_payment)

            if payment_form == words_payment[-2]:   # VISA → CR
                payment_form = 'CR'
            if payment_form is None or payment_form == words_payment[-1] or payment_form == 'DÉB':
                payment_form = 'DEB'

            if transaction[0][2] == 'DEVOLUÇÃO':
                payment_form     = 'Devolução'
                transaction_type = 'Devolução'

            item_for_convert_to_discount = [
                transaction[0][0],
                payment_form,
                transaction_type,
                transaction[0][-1][:-1],
                'Pagamento',
                '', '', '', '',
            ]

            if len(transaction) > 2:
                item_for_convert_to_discount[-1] = ' '.join(transaction[-2])

            if transaction[0][2] == 'DEVOLUÇÃO' and self._verify_pix_is_cpf(transaction):
                item_for_convert_to_discount[5] = self._extract_cpf(transaction)

            return item_for_convert_to_discount
        else:
            pass

    def create_credit_entrace(self, transaction):
        if self._verify_credit_entrance(transaction):
            words_payment = ['DEB', 'DÉB', 'Deb']
            payment_form = self._verify_payment_form(transaction[1], words_payment)
            payment_form = 'DEB' if payment_form is not None else 'CR'
            item_for_convert = [
                transaction[0][0], payment_form,
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                'Recebimento', '', '', '',
                ' '.join(transaction[1]),
            ]
            return item_for_convert
        else:
            pass

    def create_pix_entrace(self, transaction):
        if self._verify_pix(transaction):
            item_for_convert = [
                transaction[0][0],
                transaction[0][1].split(".")[0],
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                transaction[1][0],
                '', '', '', '',
            ]
            if self._verify_if_comment_exists(transaction, 'DOC.:'):
                item_for_convert[-1] = ' '.join(transaction[-2])
            if self._verify_pix_is_cnpj(transaction):
                item_for_convert[7] = ' '.join(transaction[2])
            if self._verify_pix_is_cpf(transaction):
                if self._verify_word_key_exists(transaction, "Recebimento"):
                    item_for_convert[5] = ' '.join(transaction[2])
                    item_for_convert[6] = ' '.join(transaction[3])
                else:
                    item_for_convert[6] = ' '.join(transaction[2])
            return item_for_convert
        else:
            pass

    def create_transf_entrace(self, transaction):
        if self._verify_transf_pix(transaction):
            item_for_convert = [
                transaction[0][0], 'TED',
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                'Recebimento',
                ''.join(transaction[-2]),
                ' '.join(transaction[3]),
                '', '',
            ]
            return item_for_convert
        else:
            pass

    def create_ted_entrace(self, transaction):
        if self._verify_ted(transaction):
            item_for_convert = [
                transaction[0][0], 'TED',
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                'Recebimento', '', '', '', '',
            ]
            if self._verify_pix_is_cnpj(transaction):
                item_for_convert[7] = ''.join(transaction[2])
            return item_for_convert
        else:
            pass

    def create_dep_entrace(self, transaction):
        if self._verify_dep(transaction):
            item_for_convert = [
                transaction[0][0], 'DEPOSITO',
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                'Recebimento', '', '', '', '',
            ]
            return item_for_convert
        else:
            pass

    def create_transf_sicoob(self, transaction):
        """Transferência PIX Sicoob (saída). Forma = TRANSFERÊNCIA."""
        if self._verify_transf_sicoob(transaction):
            obs = ' '.join(transaction[-1]) if len(transaction) > 3 else ''
            return [
                transaction[0][0],
                'TRANSFERÊNCIA',
                ' '.join(transaction[0][1:-1]),
                transaction[0][-1][:-1],
                'Pagamento',
                '', '', '',
                obs,
            ]
        return None

    # ── verificações ──────────────────────────────────────────────────────────

    def _verify_pix(self, transaction):
        for sub_array in transaction:
            has_pix = any('PIX' in x or '.OUTR' in x or 'Pix' in x for x in sub_array)
            has_dev = any('DEVOLUÇÃO' in x for x in sub_array)
            if has_pix and not has_dev:
                return True
        return False

    def _verify_discount(self, transaction):
        if len(transaction) <= 2:
            return True
        if len(transaction) <= 4:
            first_item = transaction[0]
            if (first_item[1].startswith('DÉB') or
                    first_item[1].startswith('DEB') or
                    first_item[1].startswith('COMP')):
                return True
            if not first_item[1].startswith('CR'):
                return True
        return False

    def _verify_credit_entrance(self, transaction):
        first_item  = transaction[0]
        second_item = transaction[1]
        if len(transaction) >= 2:
            if 'CR' in first_item and ('SIPAG' in second_item[0] or 'CIELO' in second_item[0]):
                return True
        return False

    def _verify_pix_is_cpf(self, transaction):
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        return bool(re.compile(r'\*\*\*\.(\d{3}\.\d{3}-\*\*)').search(transaction_str))

    def _verify_pix_is_cnpj(self, transaction):
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        return bool(re.compile(r'\b\d{2}\.\d{3}\.\d{3}\d{4}-\d{2}\b').search(transaction_str))

    def _extract_cpf(self, transaction) -> str:
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        match = re.search(r'\*\*\*\.\d{3}\.\d{3}-\*\*', transaction_str)
        return match.group(0) if match else ''

    def _verify_if_comment_exists(self, transaction, word_key):
        if len(transaction) >= 3:
            antepenultimate_item = ''.join(transaction[-3])
            last_item = transaction[-1]
            if word_key in last_item:
                if (self._verify_pix_is_cpf([antepenultimate_item]) or
                        self._verify_pix_is_cnpj([antepenultimate_item])):
                    return True
        return False

    def _verify_ted(self, transaction):
        if len(transaction) >= 4:
            first_item    = ''.join(transaction[0])
            code_ted_item = ''.join(transaction[-1])
            return 'TED' in first_item or 'TED' in code_ted_item
        return False

    def _verify_word_key_exists(self, transaction, word):
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        return word in transaction_str

    def _verify_payment_form(self, transaction, words):
        transaction_str = ''.join(''.join(sublist) for sublist in transaction)
        for word in words:
            if word in transaction_str:
                return word
        return None

    def _verify_transf_pix(self, transaction):
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        return 'Transferência' in transaction_str and "REM.:" in transaction[1][0]

    def _verify_dep(self, transaction):
        transaction_str = ''.join(''.join(sublist) + ' ' for sublist in transaction)
        return 'DEP' in transaction_str and len(transaction) <= 3

    def _verify_transf_sicoob(self, transaction) -> bool:
        """TRANSF. PIX SICOOB — saída entre contas Sicoob (ex: TRANSF. FAV.:)."""
        if len(transaction[0]) >= 2 and transaction[0][1] == 'TRANSF.':
            tx_str = ''.join(''.join(sub) + ' ' for sub in transaction)
            if 'FAV.:' in tx_str or 'SICOOB' in tx_str:
                return True
        return False