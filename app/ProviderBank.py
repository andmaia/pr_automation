from abc import ABC, abstractmethod

class ProviderBank(ABC):
    @abstractmethod  
    def _verify_pix(self, transaction):
        pass
    
    @abstractmethod
    def create_discount(self, transaction):
        pass
    @abstractmethod
    def create_credit_entrace(self, transaction):
        pass
    @abstractmethod
    def create_pix_entrace(self, transaction):
        pass
    @abstractmethod
    def create_ted_entrace(self, transaction):
        pass
    
    @abstractmethod
    def create_transf_entrace(self,transaction):
        pass

    @abstractmethod
    def create_dep_entrace(self,transaction):
        pass

    @abstractmethod
    def _verify_discount(self, transaction):
        pass
    
    @abstractmethod
    def _verify_credit_entrance(self, transaction):
        pass
    
    @abstractmethod
    def _verify_pix_is_cpf(self, transaction):
        pass
    
    @abstractmethod
    def _verify_pix_is_cnpj(self, transaction):
        pass
    
    @abstractmethod
    def _verify_if_comment_exists(self, transaction):
        pass

    @abstractmethod
    def _verify_ted(self, transaction):
        pass
    @abstractmethod
    def _verify_dep(self,transaction):
        pass
    
    @abstractmethod
    def _verify_transf_pix(self,transaction):
        pass