from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import List, Dict, Any, Optional
from google_auth import GoogleAuth
import json
import logging
import time
import random
import string

logger = logging.getLogger(__name__)

def generate_unique_id():
    """Generate a unique ID using timestamp and random string."""
    timestamp = int(time.time() * 1000)  # Current timestamp in milliseconds
    random_str = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    return f'slide_{timestamp}_{random_str}'

class GoogleSlides:
    def __init__(self, auth: GoogleAuth):
        self.service = build('slides', 'v1', credentials=auth.get_credentials())
        
    def create_presentation(self, title: str) -> str:
        """Create a new Google Slides presentation."""
        try:
            presentation = {
                'title': title
            }
            presentation = self.service.presentations().create(body=presentation).execute()
            return presentation.get('presentationId')
        except HttpError as error:
            logger.error(f'An error occurred: {error}')
            return None

    def add_slide(self, presentation_id: str, title: str, content: str) -> Optional[Dict[str, Any]]:
        """Add a new slide to the presentation."""
        try:
            logger.info(f"Starting to add slide with title: {title}")
            logger.info(f"Content: {content}")
            logger.info(f"Presentation ID: {presentation_id}")
            
            # First, verify the presentation exists and we have access
            try:
                presentation = self.service.presentations().get(
                    presentationId=presentation_id
                ).execute()
                logger.info(f"Successfully accessed presentation: {presentation.get('title')}")
            except HttpError as error:
                logger.error(f"Failed to access presentation: {error}")
                logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                return None

            # Generate a unique objectId for the new slide
            unique_id = generate_unique_id()
            logger.info(f"Generated unique slide ID: {unique_id}")

            # Create a new slide with title and body using the unique ID
            requests = [
                {
                    'createSlide': {
                        'objectId': unique_id,
                        'insertionIndex': 1,
                        'slideLayoutReference': {
                            'predefinedLayout': 'TITLE_AND_BODY'
                        }
                    }
                }
            ]
            
            # Create the slide first
            logger.info(f"Creating slide with title: {title}")
            try:
                response = self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                logger.info("Successfully created slide")
            except HttpError as error:
                logger.error(f"Failed to create slide: {error}")
                logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                return None

            # Get the created slide ID
            slide_id = response.get('replies')[0].get('createSlide').get('objectId')
            logger.info(f"Created slide with ID: {slide_id}")

            # Get the slide to find the title and body shape IDs
            logger.info("Fetching slide details")
            try:
                slide = self.service.presentations().pages().get(
                    presentationId=presentation_id,
                    pageObjectId=slide_id
                ).execute()
                logger.info("Successfully fetched slide details")
            except HttpError as error:
                logger.error(f"Failed to fetch slide details: {error}")
                logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                return None

            # Find the title and body shape IDs
            title_shape_id = None
            body_shape_id = None
            
            for element in slide.get('pageElements', []):
                if 'shape' in element:
                    shape = element['shape']
                    if 'placeholder' in shape:
                        placeholder = shape['placeholder']
                        if placeholder['type'] == 'TITLE':
                            title_shape_id = element['objectId']
                            logger.info(f"Found title shape ID: {title_shape_id}")
                        elif placeholder['type'] == 'BODY':
                            body_shape_id = element['objectId']
                            logger.info(f"Found body shape ID: {body_shape_id}")

            if not title_shape_id or not body_shape_id:
                logger.error("Could not find title or body shape IDs")
                logger.error(f"Title shape ID: {title_shape_id}")
                logger.error(f"Body shape ID: {body_shape_id}")
                return None

            # Update the text content with original Korean text and spaces
            requests = [
                {
                    'insertText': {
                        'objectId': title_shape_id,
                        'text': title  # Use original Korean title
                    }
                },
                {
                    'insertText': {
                        'objectId': body_shape_id,
                        'text': content  # Use original Korean content
                    }
                }
            ]
            
            logger.info("Updating slide content")
            try:
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                logger.info("Successfully updated slide content")
            except HttpError as error:
                logger.error(f"Failed to update slide content: {error}")
                logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                return None
            
            # Get slide dimensions
            dimensions = self.get_slide_dimensions(presentation_id, slide_id)
            
            logger.info("Successfully added slide")
            return {
                'success': True,
                'slide_id': slide_id,
                'dimensions': dimensions
            }
        except HttpError as error:
            logger.error(f'An error occurred while adding slide: {error}')
            logger.error(f'Error details: {error.content.decode() if hasattr(error, "content") else "No additional details"}')
            return None
        except Exception as e:
            logger.error(f'Unexpected error occurred: {str(e)}')
            logger.error(f'Error type: {type(e)}')
            import traceback
            logger.error(f'Traceback: {traceback.format_exc()}')
            return None

    def add_image(self, presentation_id: str, slide_id: str, image_url: str,
                 x: float = 100, y: float = 100,
                 width: float = 400, height: float = 300,
                 rotation: float = 0.0) -> Optional[str]:
        """Add an image to a specific slide.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to add the image to
            image_url (str): URL of the image to add
            x (float, optional): X position in points (default: 100)
            y (float, optional): Y position in points (default: 100)
            width (float, optional): Width in points (default: 400)
            height (float, optional): Height in points (default: 300)
            rotation (float, optional): Rotation angle in degrees (default: 0)
            
        Returns:
            str: Object ID of the created image, or None if failed
        """
        try:
            # Generate unique ID for the image
            image_id = generate_unique_id()
            
            # Create image request
            requests = [{
                'createImage': {
                    'objectId': image_id,
                    'url': image_url,
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': {'magnitude': width, 'unit': 'PT'},
                            'height': {'magnitude': height, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': x,
                            'translateY': y,
                            'unit': 'PT'
                        }
                    }
                }
            }]
            
            # Add rotation if specified
            if rotation != 0:
                requests.append({
                    'updatePageElementTransform': {
                        'objectId': image_id,
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': x,
                            'translateY': y,
                            'unit': 'PT',
                            'rotation': rotation
                        },
                        'applyMode': 'RELATIVE'
                    }
                })
            
            # Execute the requests
            response = self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            
            return image_id
        except HttpError as error:
            logger.error(f"Failed to add image: {error}")
            logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
            return None

    def get_presentation(self, presentation_id: str) -> Optional[Dict[str, Any]]:
        """Get presentation details."""
        try:
            presentation = self.service.presentations().get(
                presentationId=presentation_id
            ).execute()
            return presentation
        except HttpError as error:
            print(f'An error occurred: {error}')
            return None

    def delete_presentation(self, presentation_id: str) -> bool:
        """Delete a presentation."""
        try:
            self.service.presentations().delete(
                presentationId=presentation_id
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def search_elements(self, presentation_id: str, slide_id: str = None, element_type: str = None) -> List[Dict[str, Any]]:
        """Search for elements in the presentation.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str, optional): Specific slide ID to search in. If None, searches all slides.
            element_type (str, optional): Type of elements to search for ('shape', 'text', etc.)
        
        Returns:
            List[Dict[str, Any]]: List of found elements with their details
        """
        try:
            # Get all slides in the presentation
            presentation = self.service.presentations().get(
                presentationId=presentation_id
            ).execute()
            
            elements = []
            slides = presentation.get('slides', [])
            
            # If slide_id is specified, only search in that slide
            if slide_id:
                slides = [slide for slide in slides if slide.get('objectId') == slide_id]
            
            for slide in slides:
                slide_id = slide.get('objectId')
                logger.info(f"Searching in slide: {slide_id}")
                
                for element in slide.get('pageElements', []):
                    if element_type is None or element_type in element:
                        # Add slide_id to element info for reference
                        element['slideId'] = slide_id
                        elements.append(element)
                        logger.info(f"Found element {element.get('objectId')} in slide {slide_id}")
            
            return elements
        except HttpError as error:
            logger.error(f"Failed to search elements: {error}")
            logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
            return []

    def update_text_style(self, presentation_id: str, slide_id: str, element_id: str, 
                         font_family: str = None, font_size: float = None, 
                         font_weight: str = None, font_style: str = None,
                         foreground_color: str = None, background_color: str = None) -> bool:
        """Update text style of an element."""
        try:
            logger.info(f"Starting to update text style for element {element_id} in slide {slide_id}")
            logger.info(f"Style parameters: font_size={font_size}, font_weight={font_weight}, foreground_color={foreground_color}")
            
            # First, verify the element exists
            try:
                slide = self.service.presentations().pages().get(
                    presentationId=presentation_id,
                    pageObjectId=slide_id
                ).execute()
                logger.info(f"Successfully accessed slide: {slide_id}")
                
                # Find the element
                element_found = False
                for element in slide.get('pageElements', []):
                    if element.get('objectId') == element_id:
                        element_found = True
                        logger.info(f"Found element with ID {element_id}")
                        if 'shape' in element and 'text' in element['shape']:
                            logger.info("Element contains text content")
                        else:
                            logger.warning("Element does not contain text content")
                        break
                
                if not element_found:
                    logger.error(f"Element with ID {element_id} not found in slide")
                    return False
                    
            except HttpError as error:
                logger.error(f"Failed to access slide: {error}")
                logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                return False

            requests = []
            
            # Build style update request
            style = {}
            if font_family:
                style['fontFamily'] = font_family
            if font_size:
                style['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
            if font_weight:
                style['fontWeight'] = font_weight
            if font_style:
                style['fontStyle'] = font_style
            if foreground_color:
                style['foregroundColor'] = {'opaqueColor': {'rgbColor': self._parse_color(foreground_color)}}
            if background_color:
                style['backgroundColor'] = {'opaqueColor': {'rgbColor': self._parse_color(background_color)}}
            
            if style:
                requests.append({
                    'updateTextStyle': {
                        'objectId': element_id,
                        'textRange': {'type': 'ALL'},
                        'style': style,
                        'fields': '*'
                    }
                })
                logger.info(f"Prepared style update request: {json.dumps(requests, indent=2)}")
            
            if requests:
                try:
                    response = self.service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    logger.info("Successfully updated text style")
                    return True
                except HttpError as error:
                    logger.error(f"Failed to update text style: {error}")
                    logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
                    return False
            
            logger.warning("No style updates to apply")
            return False
        except Exception as e:
            logger.error(f'Unexpected error occurred: {str(e)}')
            logger.error(f'Error type: {type(e)}')
            import traceback
            logger.error(f'Traceback: {traceback.format_exc()}')
            return False

    def update_shape_style(self, presentation_id: str, slide_id: str, element_id: str,
                          width: float = None, height: float = None,
                          x: float = None, y: float = None,
                          fill_color: str = None, border_color: str = None,
                          border_width: float = None) -> bool:
        """Update shape style (size, position, colors, border)."""
        try:
            requests = []
            
            # Build transform update
            transform = {}
            if width or height:
                transform['size'] = {}
                if width:
                    transform['size']['width'] = {'magnitude': width, 'unit': 'PT'}
                if height:
                    transform['size']['height'] = {'magnitude': height, 'unit': 'PT'}
            if x is not None or y is not None:
                transform['transform'] = {}
                if x is not None:
                    transform['transform']['translateX'] = x
                if y is not None:
                    transform['transform']['translateY'] = y
            
            if transform:
                requests.append({
                    'updatePageElementTransform': {
                        'objectId': element_id,
                        'transform': transform,
                        'applyMode': 'RELATIVE'
                    }
                })
            
            # Build style update
            style = {}
            if fill_color:
                style['fill'] = {'solidFill': {'color': {'rgbColor': self._parse_color(fill_color)}}}
            if border_color:
                style['outline'] = {
                    'outlineFill': {'solidFill': {'color': {'rgbColor': self._parse_color(border_color)}}}
                }
            if border_width is not None:
                style['outline'] = style.get('outline', {})
                style['outline']['weight'] = {'magnitude': border_width, 'unit': 'PT'}
            
            if style:
                requests.append({
                    'updateShapeProperties': {
                        'objectId': element_id,
                        'shapeProperties': style,
                        'fields': '*'
                    }
                })
            
            if requests:
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to update shape style: {error}")
            return False

    def delete_element(self, presentation_id: str, slide_id: str, element_id: str) -> bool:
        """Delete an element from a slide."""
        try:
            requests = [{
                'deleteObject': {
                    'objectId': element_id
                }
            }]
            
            self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to delete element: {error}")
            return False

    def _parse_color(self, color: str) -> Dict[str, float]:
        """Parse color string (hex or RGB) into RGB color object."""
        if color.startswith('#'):
            color = color[1:]
        if len(color) == 6:
            r = int(color[0:2], 16) / 255.0
            g = int(color[2:4], 16) / 255.0
            b = int(color[4:6], 16) / 255.0
        else:
            raise ValueError("Invalid color format. Use hex (#RRGGBB) or RGB (r,g,b)")
        return {'red': r, 'green': g, 'blue': b}

    def add_shape(self, presentation_id: str, slide_id: str, 
                 shape_type: str, x: float, y: float,
                 width: float, height: float,
                 fill_color: str = None, border_color: str = None,
                 border_width: float = None) -> Optional[str]:
        """Add a shape to a slide.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to add the shape to
            shape_type (str): Type of shape ('RECTANGLE', 'TRIANGLE', 'ELLIPSE', etc.)
            x (float): X position in points
            y (float): Y position in points
            width (float): Width in points
            height (float): Height in points
            fill_color (str, optional): Fill color in hex format (e.g., '#FF0000')
            border_color (str, optional): Border color in hex format (e.g., '#000000')
            border_width (float, optional): Border width in points
            
        Returns:
            str: Object ID of the created shape, or None if failed
        """
        try:
            # Generate unique ID for the shape
            shape_id = generate_unique_id()
            
            # Create shape request
            requests = [{
                'createShape': {
                    'objectId': shape_id,
                    'shapeType': shape_type,
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': {'magnitude': width, 'unit': 'PT'},
                            'height': {'magnitude': height, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': x,
                            'translateY': y,
                            'unit': 'PT'
                        }
                    }
                }
            }]
            
            # Add style properties if specified
            style = {}
            if fill_color:
                style['fill'] = {'solidFill': {'color': {'rgbColor': self._parse_color(fill_color)}}}
            if border_color:
                style['outline'] = {
                    'outlineFill': {'solidFill': {'color': {'rgbColor': self._parse_color(border_color)}}}
                }
            if border_width is not None:
                style['outline'] = style.get('outline', {})
                style['outline']['weight'] = {'magnitude': border_width, 'unit': 'PT'}
            
            if style:
                requests.append({
                    'updateShapeProperties': {
                        'objectId': shape_id,
                        'shapeProperties': style,
                        'fields': '*'
                    }
                })
            
            # Execute the requests
            response = self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            
            return shape_id
        except HttpError as error:
            logger.error(f"Failed to add shape: {error}")
            logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
            return None

    def add_line(self, presentation_id: str, slide_id: str,
                start_x: float, start_y: float,
                end_x: float, end_y: float,
                line_color: str = '#000000',
                line_width: float = 1.0,
                line_type: str = 'STRAIGHT') -> Optional[str]:
        """Add a line to a slide.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to add the line to
            start_x (float): Starting X position in points
            start_y (float): Starting Y position in points
            end_x (float): Ending X position in points
            end_y (float): Ending Y position in points
            line_color (str, optional): Line color in hex format (e.g., '#000000')
            line_width (float, optional): Line width in points
            line_type (str, optional): Type of line ('STRAIGHT', 'CURVED', 'ELBOW', 'BENT')
            
        Returns:
            str: Object ID of the created line, or None if failed
        """
        try:
            # Generate unique ID for the line
            line_id = generate_unique_id()
            
            # Calculate line dimensions
            width = abs(end_x - start_x)
            height = abs(end_y - start_y)
            
            # Create line request
            requests = [{
                'createLine': {
                    'objectId': line_id,
                    'lineCategory': line_type,
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': {'magnitude': width, 'unit': 'PT'},
                            'height': {'magnitude': height, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': min(start_x, end_x),
                            'translateY': min(start_y, end_y),
                            'unit': 'PT'
                        }
                    }
                }
            }]
            
            # Add line style
            style = {
                'outline': {
                    'outlineFill': {'solidFill': {'color': {'rgbColor': self._parse_color(line_color)}}},
                    'weight': {'magnitude': line_width, 'unit': 'PT'}
                }
            }
            
            requests.append({
                'updateShapeProperties': {
                    'objectId': line_id,
                    'shapeProperties': style,
                    'fields': '*'
                }
            })
            
            # Execute the requests
            response = self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            
            return line_id
        except HttpError as error:
            logger.error(f"Failed to add line: {error}")
            logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
            return None

    def get_slide_dimensions(self, presentation_id: str, slide_id: str) -> Optional[Dict[str, float]]:
        """Get the dimensions of a slide.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to get dimensions for
            
        Returns:
            Dict[str, float]: Dictionary containing width and height in points, or None if failed
        """
        try:
            # Get the slide details
            slide = self.service.presentations().pages().get(
                presentationId=presentation_id,
                pageObjectId=slide_id
            ).execute()
            
            # Get the page size
            page_size = slide.get('pageProperties', {}).get('pageSize', {})
            
            # Convert to points (1 inch = 96 points)
            width = page_size.get('width', {}).get('magnitude', 960)  # Default to 10 inches
            height = page_size.get('height', {}).get('magnitude', 540)  # Default to 5.625 inches
            
            return {
                'width': width,
                'height': height,
                'width_inches': width / 96,
                'height_inches': height / 96
            }
        except HttpError as error:
            logger.error(f"Failed to get slide dimensions: {error}")
            logger.error(f"Error details: {error.content.decode() if hasattr(error, 'content') else 'No additional details'}")
            return None

    def update_slide_background(self, presentation_id: str, slide_id: str,
                              background_color: str = None,
                              background_image_url: str = None) -> bool:
        """Update slide background with color or image.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to update
            background_color (str, optional): Background color in hex format (e.g., '#FFFFFF')
            background_image_url (str, optional): URL of the background image
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            requests = []
            
            if background_color:
                requests.append({
                    'updatePageProperties': {
                        'objectId': slide_id,
                        'pageProperties': {
                            'colorScheme': {
                                'background': {
                                    'opaqueColor': {
                                        'rgbColor': self._parse_color(background_color)
                                    }
                                }
                            }
                        },
                        'fields': 'colorScheme.background'
                    }
                })
            
            if background_image_url:
                requests.append({
                    'updatePageProperties': {
                        'objectId': slide_id,
                        'pageProperties': {
                            'background': {
                                'image': {
                                    'imageUrl': background_image_url
                                }
                            }
                        },
                        'fields': 'background.image'
                    }
                })
            
            if requests:
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to update slide background: {error}")
            return False

    def update_slide_layout(self, presentation_id: str, slide_id: str,
                          layout_type: str) -> bool:
        """Update slide layout.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to update
            layout_type (str): Type of layout ('TITLE', 'TITLE_AND_BODY', 'MAIN_POINT', etc.)
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            requests = [{
                'updateSlideLayoutProperties': {
                    'objectId': slide_id,
                    'slideLayoutProperties': {
                        'predefinedLayout': layout_type
                    },
                    'fields': 'predefinedLayout'
                }
            }]
            
            self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to update slide layout: {error}")
            return False

    def update_slide_transition(self, presentation_id: str, slide_id: str,
                              transition_type: str = 'FADE',
                              duration: str = 'SLOW') -> bool:
        """Update slide transition effect.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to update
            transition_type (str): Type of transition ('FADE', 'SLIDE', 'ZOOM', etc.)
            duration (str): Duration of transition ('SLOW', 'MEDIUM', 'FAST')
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            requests = [{
                'updatePageProperties': {
                    'objectId': slide_id,
                    'pageProperties': {
                        'transition': {
                            'type': transition_type,
                            'duration': duration
                        }
                    },
                    'fields': 'transition'
                }
            }]
            
            self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to update slide transition: {error}")
            return False

    def add_slide_notes(self, presentation_id: str, slide_id: str,
                       notes_text: str) -> bool:
        """Add or update speaker notes for a slide.
        
        Args:
            presentation_id (str): Google Slides presentation ID
            slide_id (str): ID of the slide to update
            notes_text (str): Text for speaker notes
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # First, get the notes page ID
            slide = self.service.presentations().pages().get(
                presentationId=presentation_id,
                pageObjectId=slide_id
            ).execute()
            
            notes_page_id = slide.get('notesPage', {}).get('notesId')
            if not notes_page_id:
                logger.error("Could not find notes page ID")
                return False
            
            # Find the notes text shape
            notes_page = self.service.presentations().pages().get(
                presentationId=presentation_id,
                pageObjectId=notes_page_id
            ).execute()
            
            notes_shape_id = None
            for element in notes_page.get('pageElements', []):
                if 'shape' in element and 'text' in element['shape']:
                    notes_shape_id = element['objectId']
                    break
            
            if not notes_shape_id:
                logger.error("Could not find notes text shape")
                return False
            
            # Update the notes text
            requests = [{
                'insertText': {
                    'objectId': notes_shape_id,
                    'text': notes_text
                }
            }]
            
            self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            return True
        except HttpError as error:
            logger.error(f"Failed to update slide notes: {error}")
            return False 