# Deploy na Oracle Cloud (Always Free)

## 1. Criar a VM
Ver instruções na conversa com o Claude — resumo: Compute → Create Instance →
shape `VM.Standard.A1.Flex` (ARM, Always Free) → imagem Ubuntu → gerar/baixar
chave SSH → Create.

## 2. Conectar via SSH
No seu computador (PowerShell), com a chave baixada:

```
ssh -i "caminho\para\sua-chave.key" ubuntu@<IP-PUBLICO-DA-VM>
```

## 3. Rodar o setup
Já dentro da VM (via SSH):

```bash
curl -O https://raw.githubusercontent.com/gilmariolima/checker-full/main/deploy/setup.sh
bash setup.sh
```

O script clona o repositório, instala as dependências, cria o serviço
`checker-full` (systemd) e já deixa rodando.

## 4. Abrir a porta na Oracle Cloud Console
Isso é feito no navegador, não no terminal:

`Networking → Virtual Cloud Networks → (sua VCN) → Security Lists → Default
Security List → Add Ingress Rules`

- Source CIDR: `0.0.0.0/0`
- Destination Port Range: `8000`
- Protocol: `TCP`

## 5. Acessar
`http://<IP-PUBLICO-DA-VM>:8000`

## Atualizar depois de um novo `git push`
Via SSH na VM:

```bash
cd ~/checker-full
git pull origin main
./.venv/bin/pip install -r requirements.txt
sudo systemctl restart checker-full
```

## Comandos úteis
```bash
sudo systemctl status checker-full   # ver se está rodando
sudo journalctl -u checker-full -f   # ver logs em tempo real
sudo systemctl restart checker-full  # reiniciar
```
