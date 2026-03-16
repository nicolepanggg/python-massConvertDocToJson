# bulk_docx_to_single_json_simplified.py
from fileinput import filename
from pathlib import Path
import json
import shutil
import re
from docx import Document
from docx.oxml.ns import qn
import mimetypes

source_folder = Path(r"D:\2026Project\python-massConvertDocToJson\docx")
output_json   = Path("paragraphs.json")
images_folder = Path("images/stories")
thumb_folder  = Path("images/stories/thumb")
images_folder.mkdir(parents=True, exist_ok=True)
thumb_folder.mkdir(parents=True, exist_ok=True)

all_documents = []
success_count = 0
fail_count    = 0
story_counter = 1

def remove_s_suffix(text):
    """Remove _S or _s from text"""
    return re.sub(r'_[Ss]', '', text)

def is_heading_style(style_name):
    name_lower = style_name.lower()
    return (
        "heading" in name_lower or
        "title" in name_lower or
        style_name.startswith("標題") or
        style_name in ["Heading 1", "Heading 2", "Heading 3", "1.", "1.1", "Section"]
    )

def extract_images_from_para(para, doc, safe_stem, img_counter):
    img_tags = []
    for blip in para._element.iter(qn("a:blip")):
        rId = blip.get(qn("r:embed"))
        if not rId:
            continue
        try:
            image_part = doc.part.related_parts[rId]
            img_bytes  = image_part.blob

            content_type = image_part.content_type
            ext = mimetypes.guess_extension(content_type) or ".png"
            if ext in (".jpe", ".jpeg"):
                ext = ".jpg"
            if not ext.startswith("."):
                ext = "." + ext

            img_filename = f"{safe_stem}_圖片{img_counter}{ext}"
            (images_folder / img_filename).write_bytes(img_bytes)
            img_tags.append((f'<div><img src="/images/stories/{img_filename}"/></div>', images_folder / img_filename))
            img_counter += 1
        except Exception:
            continue

    return img_tags, img_counter


for docx_file in sorted(source_folder.rglob("*.docx")):
    try:
        doc = Document(docx_file)

        safe_stem = (
            docx_file.stem
            .replace(" ", "_")
            .replace("─", "-")
            .replace("：", "_")
        )
        img_index             = 1
        current_group         = None
        pending_content_parts = []
        section_count         = 0
        doc_first_image_path  = None

        def flush_group(group):
            global story_counter, doc_first_image_path

            thumb_path = None
            if doc_first_image_path and doc_first_image_path.exists():
                ext        = doc_first_image_path.suffix
                thumb_name = f"story_{story_counter:02d}{ext}"
                shutil.copy2(doc_first_image_path, thumb_folder / thumb_name)
                thumb_path = f"/images/stories/thumb/{thumb_name}"
                story_counter += 1

            all_documents.append({
                "filename": remove_s_suffix(docx_file.name),
                "icon":    thumb_path,
                "title":    remove_s_suffix(group["title"]),
                "description": "",
                "button":   "",
                "content":  group["content"]
            })

        for para in doc.paragraphs:
            text = para.text.strip()

            # 忽略純底線分隔線（任意長度）
            if text and all(c == '_' for c in text):
                continue

            img_tags, img_index = extract_images_from_para(para, doc, safe_stem, img_index)

            # 記錄文件第一張圖片（只記一次）
            if doc_first_image_path is None:
                for _, img_path in img_tags:
                    doc_first_image_path = img_path
                    break

            img_tag_strings = [tag for tag, _ in img_tags]

            is_heading = bool(text) and (
                is_heading_style(para.style.name) or
                text.startswith(("1.", "2.", "3.", "一.", "二."))
            )

            if is_heading:
                if current_group is not None:
                    flush_group(current_group)
                    section_count += 1
                elif pending_content_parts:
                    thumb_path = None
                    if doc_first_image_path and doc_first_image_path.exists():
                        ext        = doc_first_image_path.suffix
                        thumb_name = f"story_{story_counter:02d}{ext}"
                        shutil.copy2(doc_first_image_path, thumb_folder / thumb_name)
                        thumb_path = f"/images/stories/thumb/{thumb_name}"
                        story_counter += 1

                    all_documents.append({
                        "filename": remove_s_suffix(docx_file.name),
                        "icon":    thumb_path,
                        "title":    remove_s_suffix(docx_file.stem),
                        "description": "",
                        "button":   "",
                        "content":  "<br/>".join(pending_content_parts)
                    })
                    section_count += 1
                    pending_content_parts = []

                current_group = {
                    "title":   text,
                    "content": "<br/>".join(img_tag_strings)
                }

            else:
                parts = []
                if text:
                    parts.append(text)
                parts.extend(img_tag_strings)
                combined = "<br/>".join(parts)

                if not combined:
                    continue

                if current_group is not None:
                    if current_group["content"]:
                        current_group["content"] += "<br/>" + combined
                    else:
                        current_group["content"] = combined
                else:
                    pending_content_parts.append(combined)

        # 最後一個群組
        if current_group is not None:
            flush_group(current_group)
            section_count += 1
        elif pending_content_parts:
            thumb_path = None
            if doc_first_image_path and doc_first_image_path.exists():
                ext        = doc_first_image_path.suffix
                thumb_name = f"story_{story_counter:02d}{ext}"
                shutil.copy2(doc_first_image_path, thumb_folder / thumb_name)
                thumb_path = f"/images/stories/thumb/{thumb_name}"
                story_counter += 1

            all_documents.append({
                "filename": remove_s_suffix(docx_file.name),
                "icon":    thumb_path,
                "title":    remove_s_suffix(docx_file.stem),
                "description": "",
                "button":   "",
                "content":  "<br/>".join(pending_content_parts)
            })
            section_count += 1

        success_count += 1
        print(f"成功 [{success_count}]: {docx_file.name}   (段落數: {section_count})")

    except Exception as e:
        fail_count += 1
        print(f"× 失敗 {docx_file.name}: {e}")

final_output = {"documents": all_documents}

with open(output_json, "w", encoding="utf-8") as f:
    json.dump(final_output, f, ensure_ascii=False, indent=2)

print(f"\n完成！成功：{success_count}　失敗：{fail_count}")
print(f"輸出檔：{output_json}")
print(f"圖片資料夾：{images_folder}")
print(f"縮圖資料夾：{thumb_folder}　(共 {story_counter - 1} 張)")
