import os
import logging
from docx import Document
import glob
import customtkinter as ctk
import threading
import time
from pathlib import Path
import sys
import win32com.client
import pythoncom
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='word_processor.log'
)
class NeonLabel(ctk.CTkLabel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.neon_colors = ["#00ff00", "#0000ff", "#ff00ff", "#00ffff"]
        self.current_color = 0
        self.animate()
    def animate(self):
        if not hasattr(self, '_is_destroyed'):
            self.configure(text_color=self.neon_colors[self.current_color])
            self.current_color = (self.current_color + 1) % len(self.neon_colors)
            self.after(1000, self.animate)
    def destroy(self):
        self._is_destroyed = True
        super().destroy()
class ProcessingAnimation(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(fg_color="transparent")
        self.dots = ["⣾", "⣽", "⣻", "⢿", "⡿", "⣟", "⣯", "⣷"]
        self.current_dot = 0
        self.label = NeonLabel(self, text="", font=("Segoe UI", 24))
        self.label.pack(pady=10)
        self.is_animating = False
    def start(self):
        self.is_animating = True
        self.animate()
    def stop(self):
        self.is_animating = False
        self.label.configure(text="")
    def animate(self):
        if self.is_animating:
            self.label.configure(text=self.dots[self.current_dot])
            self.current_dot = (self.current_dot + 1) % len(self.dots)
            self.after(100, self.animate)
class ModernWordProcessor(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Editor de Documentos Word")
        self.geometry("1200x800")
        self.grid_columnconfigure(0, weight=3)  
        self.grid_columnconfigure(1, weight=2)  
        self.grid_rowconfigure(0, weight=1)
        self.left_frame = ctk.CTkFrame(self)
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.setup_left_frame()
        self.right_frame = ctk.CTkFrame(self)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.setup_right_frame()
        self.processing = False
        self.word_folder = None
        self.pdf_folder = None
        self.word_app = None
    def setup_left_frame(self):
        title_label = ctk.CTkLabel(
            self.left_frame,
            text="Editor de Documentos Word",
            font=("Segoe UI", 24, "bold"),
            text_color="#ffffff"
        )
        title_label.pack(pady=(20, 10), padx=20)
        instructions_frame = ctk.CTkFrame(self.left_frame)
        instructions_frame.pack(fill="x", padx=20, pady=10)
        instructions_text = "Este programa permite substituir textos em documentos Word e convertê-los para PDF."
        instructions_label = ctk.CTkLabel(
            instructions_frame,
            text=instructions_text,
            font=("Segoe UI", 12),
            justify="left"
        )
        instructions_label.pack(padx=15, pady=15)
        config_label = ctk.CTkLabel(
            self.left_frame,
            text="Configurações de Substituição",
            font=("Segoe UI", 16, "bold")
        )
        config_label.pack(pady=(20, 10))
        config_frame1 = ctk.CTkFrame(self.left_frame)
        config_frame1.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(config_frame1, text="Texto Original 1", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(5,0))
        self.find_text1 = ctk.CTkEntry(config_frame1, placeholder_text="Digite o texto que deseja encontrar")
        self.find_text1.pack(fill="x", padx=10, pady=(0,5))

        ctk.CTkLabel(config_frame1, text="Novo Texto 1", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(5,0))
        self.replace_text1 = ctk.CTkEntry(config_frame1, placeholder_text="Digite o texto que substituirá o original")
        self.replace_text1.pack(fill="x", padx=10, pady=(0,5))

        config_frame2 = ctk.CTkFrame(self.left_frame)
        config_frame2.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(config_frame2, text="Texto Original 2", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(5,0))
        self.find_text2 = ctk.CTkEntry(config_frame2, placeholder_text="Digite o texto que deseja encontrar")
        self.find_text2.pack(fill="x", padx=10, pady=(0,5))

        ctk.CTkLabel(config_frame2, text="Novo Texto 2", font=("Segoe UI", 12)).pack(anchor="w", padx=10, pady=(5,0))
        self.replace_text2 = ctk.CTkEntry(config_frame2, placeholder_text="Digite o texto que substituirá o original")
        self.replace_text2.pack(fill="x", padx=10, pady=(0,5))

        self.process_button = ctk.CTkButton(
            self.left_frame,
            text="Iniciar Processamento",
            command=self.start_processing,
            font=("Segoe UI", 14, "bold"),
            height=40,
            fg_color="#1f538d",
            hover_color="#14375e"
        )
        self.process_button.pack(pady=20)
        self.progress_label = ctk.CTkLabel(
            self.left_frame,
            text="Pronto para começar!",
            font=("Segoe UI", 12)
        )
        self.progress_label.pack(pady=(10, 0))
        self.progress_bar = ctk.CTkProgressBar(self.left_frame)
        self.progress_bar.pack(fill="x", padx=20, pady=10)
        self.progress_bar.set(0)
    def setup_right_frame(self):
        log_title = ctk.CTkLabel(
            self.right_frame,
            text="logs",
            font=("Segoe UI", 16, "bold")
        )
        log_title.pack(pady=(20, 10))
        log_container = ctk.CTkFrame(self.right_frame)
        log_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = ctk.CTkTextbox(
            log_container,
            font=("Consolas", 12),
            wrap="word"
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
    def log(self, message, level="info"):
        colors = {
            "info": "#ffffff",
            "success": "#00ff00",
            "error": "#ff0000",
            "warning": "#ffff00"
        }
        timestamp = time.strftime("%H:%M:%S")
        if level == "success":
            prefix = "✓"
        elif level == "error":
            prefix = "×"
        elif level == "warning":
            prefix = "!"
        else:
            prefix = "→"
        formatted_message = f"[{timestamp}] {prefix} {message}\n"
        self.log_text.insert("end", formatted_message)
        self.log_text.see("end")
        end_index = self.log_text.index("end-1c")
        start_index = f"{float(end_index) - 1:.1f}"
        self.log_text.tag_add(level, start_index, end_index)
        self.log_text.tag_config(level, foreground=colors[level])
    def update_status(self, message, progress=None):
        self.progress_label.configure(text=message)
        if progress is not None:
            self.progress_bar.set(progress)
    def normalize_text(self, text):
        return ' '.join(text.split()).lower()
    def convert_to_pdf(self, word_path):
        def cleanup_word():
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                self.word_app = None
            try:
                pythoncom.CoUninitialize()
            except:
                pass

        def create_word_instance():
            cleanup_word()
            time.sleep(1)
            pythoncom.CoInitialize()
            self.word_app = win32com.client.DispatchEx("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = False

        pdf_path = os.path.join(
            self.pdf_folder,
            os.path.splitext(os.path.basename(word_path))[0] + '.pdf'
        )
        
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                self.log(f"PDF antigo removido: {os.path.basename(pdf_path)}", "info")
            except Exception as e:
                self.log(f"Não foi possível remover o PDF antigo: {str(e)}", "warning")
                return False

        # Primeira tentativa - método normal
        try:
            create_word_instance()
            doc = self.word_app.Documents.Open(os.path.abspath(word_path))
            doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            cleanup_word()

            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                self.log(f"PDF gerado com sucesso: {os.path.basename(pdf_path)}", "success")
                return True
        except Exception as e:
            if 'doc' in locals():
                try:
                    doc.Close()
                except:
                    pass
            cleanup_word()
            self.log("Primeira tentativa falhou, tentando método alternativo...", "warning")

        # Segunda tentativa - criar cópia e tentar novamente
        try:
            temp_doc_path = os.path.join(
                os.path.dirname(word_path),
                f"temp_{os.path.basename(word_path)}"
            )
            
            # Criar uma cópia do documento
            import shutil
            shutil.copy2(word_path, temp_doc_path)
            
            create_word_instance()
            doc = self.word_app.Documents.Open(os.path.abspath(temp_doc_path))
            doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            cleanup_word()
            
            # Remover arquivo temporário
            try:
                os.remove(temp_doc_path)
            except:
                pass

            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                self.log(f"PDF gerado com sucesso (método alternativo): {os.path.basename(pdf_path)}", "success")
                return True
            else:
                raise Exception("PDF não foi criado corretamente")

        except Exception as e:
            if 'doc' in locals():
                try:
                    doc.Close()
                except:
                    pass
            cleanup_word()
            try:
                os.remove(temp_doc_path)
            except:
                pass
            self.log(f"Erro ao gerar PDF: {str(e)}", "error")
            return False
    def process_document(self, file_path):
        try:
            self.log(f"Analisando documento: {os.path.basename(file_path)}", "info")
            doc = Document(file_path)
            modified = False
            
            find_text1 = self.find_text1.get().strip()
            replace_text1 = self.replace_text1.get().strip()
            find_text2 = "Químico Responsável: CRQ 03413608 - 3"
            replace_text2 = "Químico Responsável: UEVERTON DE SOUZA BARBOSA CRQ 04167832 - 4ª Região"

            for section in doc.sections:
                for paragraph in section.footer.paragraphs:
                    original_text = paragraph.text
                    normalized_text = self.normalize_text(original_text)
                    
                    if find_text1 and self.normalize_text(find_text1) in normalized_text:
                        paragraph.text = replace_text1
                        modified = True
                        self.log(f"Texto atualizado: {find_text1} → {replace_text1}", "success")
                    
                    if find_text2.lower() in normalized_text:
                        paragraph.text = replace_text2
                        modified = True
                        self.log(f"Informações do químico atualizadas em: {os.path.basename(file_path)}", "success")

            if modified:
                doc.save(file_path)
            else:
                self.log(f"Documento já está com as informações atualizadas: {os.path.basename(file_path)}", "info")
            
            return modified

        except Exception as e:
            self.log(f"Erro ao processar o documento: {str(e)}", "error")
            return False
    def process_files(self):
        try:
            if not self.find_text1.get().strip() or not self.replace_text1.get().strip():
                self.log("Por favor, preencha o primeiro campo de texto", "warning")
                self.update_status("Configuração incompleta")
                return
            self.word_folder = ctk.filedialog.askdirectory(
                title="1. Selecione a pasta com seus arquivos Word"
            )
            if not self.word_folder:
                self.update_status("Operação cancelada")
                return
            self.pdf_folder = ctk.filedialog.askdirectory(
                title="2. Selecione onde deseja salvar os PDFs"
            )
            if not self.pdf_folder:
                self.update_status("Operação cancelada")
                return
            if os.path.normpath(self.word_folder) == os.path.normpath(self.pdf_folder):
                self.log("Por favor, escolha pastas diferentes para os arquivos Word e PDF", "warning")
                self.update_status("Mesma pasta selecionada")
                return
            word_files = glob.glob(os.path.join(self.word_folder, "*.docx"))
            if not word_files:
                self.update_status("Nenhum arquivo Word encontrado")
                self.log("Não encontrei nenhum arquivo .docx na pasta selecionada", "warning")
                return
            total_files = len(word_files)
            successful_word = 0
            successful_pdf = 0
            already_updated = 0
            no_changes_needed = 0
            self.process_button.configure(state="disabled")
            self.log(f"Começando a processar {total_files} arquivos...", "info")
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                self.word_app = None
                pythoncom.CoUninitialize()
                time.sleep(1)
            for i, file in enumerate(word_files, 1):
                if not self.processing:
                    break
                progress = i / total_files
                self.update_status(f"Processando arquivo {i} de {total_files}", progress)
                was_modified = self.process_document(file)
                if was_modified:
                    successful_word += 1
                    if self.convert_to_pdf(file):
                        successful_pdf += 1
                    else:
                        self.log(f"Não consegui gerar o PDF para: {os.path.basename(file)}", "error")
                else:
                    if "já está atualizado" in self.log_text.get("1.0", "end"):
                        already_updated += 1
                    else:
                        no_changes_needed += 1
                time.sleep(0.5)
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                self.word_app = None
                pythoncom.CoUninitialize()
            self.process_button.configure(state="normal")
            final_message = (
                f"Tudo pronto!\n"
                f"• Documentos atualizados: {successful_word}/{total_files}\n"
                f"• PDFs gerados: {successful_pdf}/{total_files}\n"
                f"• Já estavam atualizados: {already_updated}/{total_files}\n"
                f"• Sem necessidade de mudanças: {no_changes_needed}/{total_files}"
            )
            self.update_status(final_message, 1.0)
            self.log(final_message, "success")
        except Exception as e:
            self.log(f"Ops! Algo deu errado: {str(e)}", "error")
            self.update_status("Erro no processamento")
            if self.word_app:
                try:
                    self.word_app.Quit()
                except:
                    pass
                self.word_app = None
                pythoncom.CoUninitialize()
            self.process_button.configure(state="normal")
    def start_processing(self):
        self.processing = True
        self.log_text.delete("1.0", "end")
        self.log("Iniciando o processamento dos arquivos...", "info")
        threading.Thread(target=self.process_files, daemon=True).start()
if __name__ == "__main__":
    app = ModernWordProcessor()
    app.mainloop() 